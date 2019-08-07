using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Exchange.WebServices.Data;
using System.DirectoryServices.AccountManagement;
using System.Threading;
using Rtf2Html;
using System.IO;

namespace MailMerger_V3
{
    public partial class main : Form
    {
        //Main application setup
        Excel.Application excel = new Excel.Application();
        OpenFileDialog openBook = new OpenFileDialog();
        int forenameColumn = -1, referenceColumn = -1, emailColumn = -1, currentRecipient = 0;
        Excel.Workbook importBook;
        Excel.Worksheet importSheet;
        Excel.Range usedRange;
        string[] columns;
        string[][] allData { get; set; }
        bool importSuccess = false;
        bool initialiseSuccess { get; set; } = false;
        string errors = "";

        //EWS stuff
        ExchangeService service = new ExchangeService();
        bool autoDiscoverFinished = false;
        Thread startAutoDiscover;
        EmailAddress sendFrom;
        FolderId sentBoxSentItems;
        FolderId sentBoxDrafts;

        //Related to files and attachments
        List<string> toDelete;
        OpenFileDialog getAttachment = new OpenFileDialog();

        //Settings
        string rtfSignature;
        string htmlSignature;
        string htmlEmailBody;
        string[] referenceAliases;
        string defaultFont;
        bool signatureContainsSignoff = false;

        public main()
        {
            InitializeComponent();
            logBox.SelectAll();
            logBox.SelectionAlignment = HorizontalAlignment.Center;
            //EWS setup
            service.UseDefaultCredentials = true;
            startAutoDiscover = new Thread(autoDiscoverThread);
            startAutoDiscover.IsBackground = true;
            startAutoDiscover.Start();
            //Import settings etc
            inboxes.Items.Add(getEmail());
            if (File.Exists("Settings.txt"))
            {
                importSettings();
                initialiseSuccess = true;
                bodyBox.Font = new Font(UsefulTools.getFontValue(defaultFont, "Name"), float.Parse(UsefulTools.getFontValue(defaultFont, "Size")));
            }
            else
            {
                //Start setup
                Setup setupForm = new Setup();
                setupForm.ShowDialog();
                if (setupForm.DialogResult == DialogResult.OK)
                {
                    importSettings();
                    initialiseSuccess = true;
                }
            }
        }

        private void importButton_Click(object sender, EventArgs e)
        {
            if (openBook.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    importWorkbook(openBook.FileName);
                    if (importSuccess)
                    {
                        displayRecipient(0);
                    }
                    importBook.Close();
                } catch (Microsoft.CSharp.RuntimeBinder.RuntimeBinderException)
                {
                    writeLog("Selected file is not a valid excel file.");
                }
            } else
            {
                writeLog("Please select an excel file.");
            }
        }

        private void previewButton_Click(object sender, EventArgs e)
        {
            if (importSuccess)
            {
                emailPrep();
                var previewEmail = createEmail(allData[currentRecipient], currentRecipient);
                previewEmail.Save(WellKnownFolderName.Drafts);
                PropertySet psPropSet = new PropertySet(EmailMessageSchema.MimeContent);
                EmailMessage emEmailMessage = EmailMessage.Bind(service, previewEmail.Id, psPropSet);
                FileStream fsFileStream = new FileStream("preview.eml", FileMode.Create);
                fsFileStream.Write(emEmailMessage.MimeContent.Content, 0, emEmailMessage.MimeContent.Content.Length);
                fsFileStream.Close();
                previewEmail.Delete(DeleteMode.HardDelete);
                System.Diagnostics.Process.Start("preview.eml");
                //Delete created resources
                foreach (string eachString in toDelete)
                {
                    if (File.Exists(eachString))
                    {
                        File.Delete(eachString);
                    }
                }
            } else
            {
                writeLog("Please import before previewing.");
            }
            
        }

        private void sendButton_Click(object sender, EventArgs e)
        {
            if (importSuccess == false)
            {
                writeLog("Please import data before attempting to send emails.");
            } else if (autoDiscoverFinished == false)
            {
                writeLog("System is still configuring your exchange profile - please wait for a couple of minutes and try again.");
            } else if (subjectBox.Text.Trim().Equals("") || subjectBox.Text.Trim().Equals("[Replace with subject]"))
            {
                writeLog("Please change the email subject before attempting to send.");
            } else if (bodyBox.Text.Trim().Equals("") || bodyBox.Text.Trim().Equals("[Replace with body of email]"))
            {
                writeLog("Please change the email body before attempting to send.");
            } else
            {
                if (MessageBox.Show("Are you sure you want to send?", "Confirm send", MessageBoxButtons.YesNo) == DialogResult.Yes)
                {
                    enableForm(false);
                    emailPrep();
                    //Update on progress
                    BackgroundWorker backgroundSend = new BackgroundWorker();
                    backgroundSend.DoWork += new DoWorkEventHandler(sendAll);
                    backgroundSend.RunWorkerCompleted += new RunWorkerCompletedEventHandler(sendFinished);
                    backgroundSend.RunWorkerAsync();

                }
            }
        }

        private void templateButton_Click(object sender, EventArgs e)
        {
            bodyBox.Rtf = rtfSignature;
        }

        private void settingsButton_Click(object sender, EventArgs e)
        {
            Setup setupForm = new Setup();
            setupForm.ShowDialog();
            importSettings();
        }

        private void addButton_Click(object sender, EventArgs e)
        {
            if (getAttachment.ShowDialog() == DialogResult.OK && File.Exists(getAttachment.FileName))
            {

                attachments.Items.Add(new Attachment(getAttachment.FileName));
                attachments.SelectedItem = attachments.Items[attachments.Items.Count - 1];
            } else
            {
                writeLog("Please select a valid file to attach.");
            }
            
        }

        private void removeButton_Click(object sender, EventArgs e)
        {
            if (attachments.Items.Count > 0)
            {
                attachments.Items.RemoveAt(attachments.SelectedIndex);
                if (attachments.Items.Count > 0)
                {
                    attachments.SelectedItem = attachments.Items[attachments.Items.Count - 1];
                }
            }
        }

        private void nextButton_Click(object sender, EventArgs e)
        {
            if (importSuccess && currentRecipient < allData.Length - 2)
            {
                displayRecipient(currentRecipient + 1);
            }
        }

        private void previousButton_Click(object sender, EventArgs e)
        {
            if (importSuccess && currentRecipient > 0)
            {
                displayRecipient(currentRecipient - 1);
            }
        }

        private void clearButton_Click(object sender, EventArgs e)
        {
            bodyBox.Text = "";
            subjectBox.Text = "";
        }
        private void writeLog(string toWrite)
        {
            logBox.Text = logBox.Text + toWrite + Environment.NewLine + Environment.NewLine;
            logBox.SelectAll();
            logBox.SelectionAlignment = HorizontalAlignment.Center;
            logBox.Select(logBox.Text.Length - 1, 1);
            logBox.ScrollToCaret();
        }
        private void writeErrors(string toWrite)
        {
            if (File.Exists("Errors.txt"))
            {
                File.AppendAllText("Errors.txt", "Error occured at " + DateTime.Now + ":" + Environment.NewLine + toWrite + Environment.NewLine);
            } else
            {
                File.Create("Errors.txt").Close();
                File.AppendAllText("Errors.txt", "Error occured at " + DateTime.Now + ":" + Environment.NewLine + toWrite + Environment.NewLine);
            }
        }
        private void importWorkbook(string filepath)
        {
            //Reset everything
            importSuccess = false;
            forenameColumn = -1;
            emailColumn = -1;
            referenceColumn = -1;
            dearBox.Text = "";
            refBox.Text = "";
            emailBox.Text = "";
            importBook = excel.Workbooks.Open(filepath);
            if (importBook.Worksheets.Count > 1)
            {
                string[] possibleSheets = new string[importBook.Worksheets.Count];
                for (int i = 1; i <= importBook.Worksheets.Count; ++i)
                {
                    possibleSheets[i - 1] = importBook.Worksheets[i].Name;
                }
                string selectedSheet = "";
                SheetSelect sheetSelect = new SheetSelect(possibleSheets);
                while (sheetSelect.DialogResult != DialogResult.OK)
                {
                    sheetSelect.ShowDialog();
                }
                selectedSheet = sheetSelect.selectedSheet;
                importSheet = importBook.Worksheets[selectedSheet];
            } else
            {
                importSheet = importBook.Worksheets[1];
            }
            usedRange = importSheet.UsedRange;
            columns = new string[usedRange.Columns.Count];
            allData = new string[usedRange.Rows.Count][];
            //Get all sheet data
            for (int row = 0; row < usedRange.Rows.Count; ++row)
            {
                allData[row] = new string[usedRange.Columns.Count];
                for (int column = 0; column < usedRange.Columns.Count; ++column)
                {
                    if (importSheet.Cells[row + 1, column + 1].Value == null)
                    {
                        allData[row][column] = "";
                    } else
                    {
                        allData[row][column] = importSheet.Cells[row + 1, column + 1].Value.ToString();
                    }
                }
            }
            //Get column headers
            for (int i = 0; i < allData[0].Length; ++i)
            {
                if (forenameColumn == -1)
                {
                    if (UsefulTools.findMatch(allData[0][i], new string[] { "forename", "firstname", "first name", "fore name" }))
                    {
                        forenameColumn = i;
                        writeLog("Forename is column " + (forenameColumn + 1));
                    }
                }
                if (referenceColumn == -1)
                {
                    if ((!(referenceAliases[0].Equals(""))) && UsefulTools.findMatch(allData[0][i], referenceAliases))
                    {
                        referenceColumn = i;
                        writeLog(referenceAliases[0] + " is column " + (referenceColumn + 1));
                    }
                }
                if (emailColumn == -1)
                {
                    if (UsefulTools.findMatch(allData[0][i], new string[] { "email", " e-mail" }))
                    {
                        emailColumn = i;
                        writeLog("Email is column " + (emailColumn + 1));
                    }
                }
            }
            if (forenameColumn != -1 && emailColumn != -1)
            {
                writeLog("Number of recipients: " + (allData.Length - 1));
                if (referenceColumn == -1)
                {
                    writeLog("Import was successful - but reference column was not found.");
                } else
                {
                    writeLog("Import successful!");
                }
                importSuccess = true;
                UsefulTools.setMergeFields(allData[0], insertMergeFieldToolStripMenuItem, bodyBox);
            } else if (forenameColumn == -1)
            {
                writeLog("Import failed - forename column missing.");
            } else if (emailColumn == -1)
            {
                writeLog("Import failed - email column missing.");
            }
        }

        private void main_FormClosing(object sender, FormClosingEventArgs e)
        {
            try
            {
                importBook.Close();
            } catch (Exception)
            {
                //Excel workbook already closed - do nothing!
            }
            if (File.Exists("preview.eml"))
            {
                File.Delete("preview.eml");
            }
            if (File.Exists("emailHTML.html"))
            {
                File.Delete("emailHTML.html");
            }
            excel.Quit();
            //outlook.Quit();
        }

        private void sheetDump()
        {
            //For debug purposes
            Console.WriteLine("Dumping sheet contents..");
            foreach (string[] row in allData)
            {
                foreach (string eachString in row)
                {
                    Console.WriteLine(eachString);
                }
            }
        }
        private void displayRecipient(int toDisplay)
        {
            currentRecipient = toDisplay;
            dearBox.Text = allData[toDisplay + 1][forenameColumn];
            if (referenceColumn != - 1)
            {
                refBox.Text = allData[toDisplay + 1][referenceColumn];
            }
            emailBox.Text = allData[toDisplay + 1][emailColumn];
        }
        private void autoDiscoverThread()
        {
            service.AutodiscoverUrl(getEmail());
            Console.WriteLine("Autodiscover completed. URL = " + service.Url);
            autoDiscoverFinished = true;
        }
        private EmailMessage createEmail(string[] recipientData, int recipientNumber)
        {
            EmailMessage email = new EmailMessage(service);
            int imageOccurences = UsefulTools.CountStringOccurrences(htmlEmailBody, "<img");
            int lastImgOccurence = 0;
            //Attach all images and embed them into html body
            Console.WriteLine("Image occurences = " + imageOccurences);
            for (int i = 0; i < imageOccurences; ++i)
            {
                lastImgOccurence = htmlEmailBody.IndexOf("<img", lastImgOccurence + 1);
                string imgSRC = UsefulTools.grabImgSRC(lastImgOccurence, htmlEmailBody);
                string fileLocation = imgSRC.Substring(1);
                htmlEmailBody.Replace(imgSRC, "CID:" + fileLocation.Substring(1));
                email.Attachments.AddFileAttachment(fileLocation);
                email.Attachments[i].IsInline = true;
                email.Attachments[i].ContentId = fileLocation.Substring(1);
                toDelete.Add(fileLocation);
            }
            string thisBody = "<span style=\"" + UsefulTools.grabStyle(htmlEmailBody) + "\">";
            thisBody += "Dear " + recipientData[forenameColumn] + "," + "<br/><br/>";
            thisBody += replaceMergeFields(htmlEmailBody, recipientNumber) + "<br/>";
            if (!signatureContainsSignoff)
            {
                thisBody += "Kind regards" + "<br/><br/>";
            }
            thisBody += UsefulTools.replaceStyle(htmlEmailBody, htmlSignature);
            email.Body = thisBody;
            email.Subject = subjectBox.Text;
            if (referenceColumn != -1)
            {
                email.Subject += " - " + referenceAliases[0] + ": " + recipientData[referenceColumn];
            }
            email.ToRecipients.Add(recipientData[emailColumn]);
            email.From = sendFrom;
            //Attach attachments
            foreach (Attachment eachAttachment in attachments.Items)
            {
                if (File.Exists(eachAttachment.filePath))
                {
                    email.Attachments.AddFileAttachment(eachAttachment.filePath);
                } else
                {
                    errors += ("Error for recipient on row " + recipientNumber + " (" + allData[recipientNumber][forenameColumn] + ") - attachment \"" + eachAttachment.fileName + "\" cannot be found." + Environment.NewLine);
                }
            }
            return email;
        }

        private void main_Load(object sender, EventArgs e)
        {
            if(!(initialiseSuccess))
            {
                //Disable for debugging
                this.Close();
            }
        }
        //MenuItems events
        private void copyToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Clipboard.SetData(DataFormats.Rtf, bodyBox.SelectedRtf);
        }

        private void hyperlinkToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string link = Prompt.ShowDialog("Enter link for hyperlink", "Hyperlink");
            UsefulTools.replaceHighlightedRtf("\\cf3\\ul " + bodyBox.SelectedText + " <" + link + ">\\cf7\\ulnone", bodyBox);
        }

        private void pasteToolStripMenuItem_Click(object sender, EventArgs e)
        {
            bodyBox.SelectedRtf = Clipboard.GetData(DataFormats.Rtf).ToString();
        }

        private void emailPrep()
        {
            htmlSignature = replaceSignatureFields(htmlSignature);
            UsefulTools.trimRtf(bodyBox);
            Rtf2Html.HtmlResult convertedRTF = RtfToHtmlConverter.RtfToHtml(bodyBox.Rtf);
            convertedRTF.WriteToFile("emailHTML.html");
            htmlEmailBody = convertedRTF.Html;
            toDelete = new List<string>();
            toDelete.Add("emailHTML.html");
            string sendingInbox = inboxes.SelectedItem.ToString().Trim();
            Mailbox sentBox = new Mailbox(sendingInbox);
            sentBoxSentItems = new FolderId(WellKnownFolderName.SentItems, sentBox);
            sentBoxDrafts = new FolderId(WellKnownFolderName.Drafts, sentBox);
            sendFrom = new EmailAddress(sendingInbox);
            if (htmlSignature.ToLower().Contains("kind regards") || htmlSignature.ToLower().Contains("best wishes") || htmlSignature.ToLower().Contains("all the best") || htmlSignature.ToLower().Contains("best regards") || htmlSignature.ToLower().Contains("yours sincerely"))
            {
                signatureContainsSignoff = true;
            } else
            {
                signatureContainsSignoff = false;
            }
        }
        private void sendAll(object sender, DoWorkEventArgs e)
        {
            toDelete = new List<string>();
            toDelete.Add("emailHTML.html");
            for (int i = 1; i < allData.Length; ++i)
            {
                if (allData[i][forenameColumn].Trim().Equals("") && allData[i][emailColumn].Trim().Equals(""))
                {
                    //Do nothing
                }
                else if (allData[i][forenameColumn].Trim().Equals(""))
                {
                    errors += ("Error for recipient on row " + i + " (" + allData[i][emailColumn] + ") - forename is missing." + Environment.NewLine);
                }
                else if (allData[i][emailColumn].Trim().Equals(""))
                {
                    errors += ("Error for recipient on row " + i + " (" + allData[i][forenameColumn] + ") - email address is missing." + Environment.NewLine);
                }
                else if (!(UsefulTools.isValidEmail(allData[i][emailColumn])))
                {
                    errors += ("Error for recipient on row " + i + " (" + allData[i][forenameColumn] + ") - email address is invalid." + Environment.NewLine);
                }  else
                {
                    sendEmail(allData[i], i);
                }
            }
            //Delete created resources
            foreach (string eachString in toDelete)
            {
                if (File.Exists(eachString))
                {
                    File.Delete(eachString);
                }
            }
            e.Result = errors;
        }
        private void sendEmail(String[] recipientData, int recipientNumber)
        {
            EmailMessage finalEmail = createEmail(recipientData, recipientNumber);
            finalEmail.Save(sentBoxDrafts);
            finalEmail.SendAndSaveCopy(sentBoxSentItems);
        }
        private void importSettings()
        {
            int currentInboxes = inboxes.Items.Count;
            if (currentInboxes > 1)
            {
                for (int i = 1; i < currentInboxes; ++i)
                {
                    inboxes.Items.Remove(inboxes.Items[1]);
                }
            }
            string stringAliases = "";
            SettingsImporter.importSettings(ref rtfSignature, ref inboxes, ref stringAliases, ref defaultFont);
            if (stringAliases.Trim().Equals("Replace with comma separated list of customer reference aliases") || stringAliases.Trim().Equals(""))
            {
                referenceAliases = new string[] { "" };
            } else if (stringAliases.IndexOf(',') != -1)
            {
                referenceAliases = stringAliases.Split(',').Select(p => p.Trim()).ToArray();
            } else
            {
                referenceAliases = new string[] {stringAliases};
            }
            if (bodyBox.Text.Trim().Equals("[Replace with body of email]"))
            {
                bodyBox.Font = new Font(UsefulTools.getFontValue(defaultFont, "Name"), float.Parse(UsefulTools.getFontValue(defaultFont, "Size")));
            }
            htmlSignature = RtfToHtmlConverter.RtfToHtml(rtfSignature).Html;
        }
        private string replaceMergeFields(string toReplace, int recipientNumber)
        {
            for (int i = 0; i < allData[0].Length; ++i)
            {
                //for each header
                if (toReplace.Contains("[[" + allData[0][i] + "]]"))
                {
                    toReplace = toReplace.Replace("[[" + allData[0][i] + "]]", allData[recipientNumber][i]);
                }
            }
            return toReplace;
        }
        private string replaceSignatureFields(string toReplace)
        {
            if (toReplace.Contains("[[Name]]"))
            {
                toReplace = toReplace.Replace("[[Name]]", getName());
            }
            if (toReplace.Contains("[[Email]]"))
            {
                toReplace = toReplace.Replace("[[Email]]", getEmail());
            }
            if (toReplace.Contains("[[Job Title]]"))
            {
                toReplace = toReplace.Replace("[[Job Title]]", getJob());
            }
            return toReplace;
        }
        private void enableForm(bool enabled)
        {
            foreach (Control c in this.Controls)
            {
                c.Enabled = enabled;
            }
        }
        private void sendFinished(object sender, RunWorkerCompletedEventArgs e)
        {
            enableForm(true);
            
            if (e.Result.ToString().Trim().Equals("")) {
                writeLog("Emails sent without errors!");
            } else
            {
                writeLog("Emails sent with errors:" + Environment.NewLine + Environment.NewLine + e.Result.ToString());
                writeErrors(e.Result.ToString());
            }
        }

        private string getName()
        {
            return System.DirectoryServices.AccountManagement.UserPrincipal.Current.DisplayName;
        }
        private string getEmail()
        {
            return System.DirectoryServices.AccountManagement.UserPrincipal.Current.EmailAddress;
        }
        private string getJob()
        {
            string jobTitle = "";
            while (!autoDiscoverFinished)
            {
                Thread.Sleep(100);
            }
            try {
                jobTitle = service.ResolveName(getEmail().Substring(0, getEmail().IndexOf('@')), ResolveNameSearchLocation.DirectoryOnly, true)[0].Contact.JobTitle;
            } catch (Exception)
            {
                while (jobTitle.Equals(""))
                {
                    jobTitle = Prompt.ShowDialog("Job title not found! Please enter job title.", "Job Title Required");
                }
            }
            
            return jobTitle;
        }
    }
}
