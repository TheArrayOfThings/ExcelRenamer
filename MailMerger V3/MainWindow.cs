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

        //EWS stuff
        ExchangeService service = new ExchangeService();
        bool autoDiscoverFinished = false;
        Thread startAutoDiscover;
        EmailAddress sendFrom;
        FolderId sentBoxSentItems;
        FolderId sentBoxDrafts;

        //Related to files and attachments
        List<string> toDelete;

        //Settings
        string rtfSignature;
        string htmlSignature;
        string htmlEmailBody;
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
            inboxes.Items.Add(System.DirectoryServices.AccountManagement.UserPrincipal.Current.EmailAddress);
            if (File.Exists("Settings.txt"))
            {
                importSettings();
                initialiseSuccess = true;
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
                    displayRecipient(0);
                    currentRecipient = 1;
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
            } else
            {
                if (MessageBox.Show("Are you sure you want to send?", "Confirm send", MessageBoxButtons.YesNo) == DialogResult.Yes)
                {
                    //Disable main form
                    //Update on progress
                    //Report on successes
                    Thread sendThread = new Thread(sendAll);
                    sendThread.IsBackground = true;
                    sendThread.Start();

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

        }

        private void removeButton_Click(object sender, EventArgs e)
        {
            Console.WriteLine(bodyBox.Rtf);
        }

        private void nextButton_Click(object sender, EventArgs e)
        {
            if (currentRecipient < allData.Length - 2)
            {
                displayRecipient(currentRecipient + 1);
            }
        }

        private void previousButton_Click(object sender, EventArgs e)
        {
            if (currentRecipient > 0)
            {
                displayRecipient(currentRecipient - 1);
            }
        }

        private void clearButton_Click(object sender, EventArgs e)
        {
            
        }
        private void writeLog(string toWrite)
        {
            logBox.Text = logBox.Text + toWrite + Environment.NewLine + Environment.NewLine;
            logBox.SelectAll();
            logBox.SelectionAlignment = HorizontalAlignment.Center;
        }
        private void importWorkbook(string filepath)
        {
            //Reset everything
            importSuccess = false;
            forenameColumn = -1;
            emailColumn = -1;
            referenceColumn = -1;
            importBook = excel.Workbooks.Open(filepath);
            importSheet = importBook.Worksheets[1];
            usedRange = importSheet.UsedRange;
            columns = new string[usedRange.Columns.Count];
            allData = new string[usedRange.Rows.Count][];
            //Get all sheet data
            for (int row = 0; row < usedRange.Rows.Count; ++row)
            {
                allData[row] = new string[usedRange.Columns.Count];
                for (int column = 0; column < usedRange.Columns.Count; ++column)
                {
                    allData[row][column] = importSheet.Cells[row + 1, column + 1].Value.ToString();
                }
            }
            //Get column headers
            for (int i = 0; i < allData[0].Length; ++i)
            {
                if (forenameColumn == -1)
                {
                    if (StringTools.findMatch(allData[0][i], new string[] { "forename", "firstname", "first name", "fore name" }))
                    {
                        forenameColumn = i;
                        writeLog("Forename is column " + (forenameColumn + 1));
                    }
                }
                if (referenceColumn == -1)
                {
                    if (StringTools.findMatch(allData[0][i], new string[] { "reference", " id" }))
                    {
                        referenceColumn = i;
                        writeLog("Reference is column " + (referenceColumn + 1));
                    }
                }
                if (emailColumn == -1)
                {
                    if (StringTools.findMatch(allData[0][i], new string[] { "email", " e-mail" }))
                    {
                        emailColumn = i;
                        writeLog("Email is column " + (emailColumn + 1));
                    }
                }
            }
            if (forenameColumn != -1 && emailColumn != -1)
            {
                writeLog("Number of recipients: " + (allData.Length - 1));
                writeLog("Import successful!");
                importSuccess = true;
                setMergeFields(allData[0]);
            }
        }

        private void main_FormClosing(object sender, FormClosingEventArgs e)
        {
            try
            {
                importBook.Close();
            } catch (Exception)
            {
                //Do nothing!
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
            service.AutodiscoverUrl(System.DirectoryServices.AccountManagement.UserPrincipal.Current.EmailAddress);
            Console.WriteLine("Autodiscover completed. URL = " + service.Url);
            autoDiscoverFinished = true;
        }
        private EmailMessage createEmail(string[] recipientData, int recipientNumber)
        {
            EmailMessage email = new EmailMessage(service);
            int imageOccurences = StringTools.CountStringOccurrences(htmlEmailBody, "<img");
            int lastImgOccurence = 0;
            //Attach all images and embed them into html body
            for (int i = 0; i < imageOccurences; ++i)
            {
                lastImgOccurence = htmlEmailBody.IndexOf("<img", lastImgOccurence);
                string imgSRC = StringTools.grabImgSRC(lastImgOccurence, htmlEmailBody);
                string fileLocation = imgSRC.Substring(1);
                htmlEmailBody.Replace(imgSRC, "CID:" + fileLocation.Substring(1));
                email.Attachments.AddFileAttachment(fileLocation);
                email.Attachments[i].IsInline = true;
                email.Attachments[i].ContentId = fileLocation.Substring(1);
                toDelete.Add(fileLocation);
            }
            string thisBody = "<span style=\"" + StringTools.grabStyle(htmlEmailBody) + "\">";
            thisBody += "Dear " + recipientData[forenameColumn] + "," + "<br/><br/>";
            thisBody += replaceMergeFields(htmlEmailBody, recipientNumber) + "<br/>";
            if (!signatureContainsSignoff)
            {
                thisBody += "Kind regards" + "<br/><br/>";
            }
            thisBody += StringTools.replaceStyle(htmlEmailBody, htmlSignature);
            email.Body = thisBody;
            email.Subject = subjectBox.Text;
            email.ToRecipients.Add(recipientData[emailColumn]);
            email.From = sendFrom;
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

        private void copyToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Clipboard.SetData(DataFormats.Rtf, bodyBox.SelectedRtf);
        }

        private void hyperlinkToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string link = Prompt.ShowDialog("Enter link for hyperlink", "Hyperlink");
            replaceHighlightedRtf("\\cf3\\ul " + bodyBox.SelectedText + " <" + link + ">\\cf7\\ulnone");
        }

        private void emailPrep()
        {
            StringTools.trimRtf(bodyBox);
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
        private void sendAll()
        {
            emailPrep();
            toDelete = new List<string>();
            toDelete.Add("emailHTML.html");
            string sendingInbox = inboxes.SelectedItem.ToString().Trim();
            Mailbox sentBox = new Mailbox(sendingInbox);
            sentBoxSentItems = new FolderId(WellKnownFolderName.SentItems, sentBox);
            sentBoxDrafts = new FolderId(WellKnownFolderName.Drafts, sentBox);
            sendFrom = new EmailAddress(sendingInbox);
            for (int i = 1; i < allData.Length; ++i)
            {
                sendEmail(allData[i], i);
            }
            //Delete created resources
            foreach (string eachString in toDelete)
            {
                if (File.Exists(eachString))
                {
                    File.Delete(eachString);
                }
            }
        }
        private void sendEmail(String[] recipientData, int recipientNumber)
        {
            EmailMessage finalEmail = createEmail(recipientData, recipientNumber);
            finalEmail.Save(sentBoxDrafts);
            finalEmail.SendAndSaveCopy(sentBoxSentItems);
        }
        private void replaceHighlightedRtf(string toAdd)
        {
            bodyBox.SelectedRtf = "{\\rtf1" + toAdd + "}}";
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
            SettingsImporter.importSettings(ref rtfSignature, ref inboxes);
            htmlSignature = RtfToHtmlConverter.RtfToHtml(rtfSignature).Html;
        }
        private void setMergeFields(String[] toSet)
        {
            for (int i = 0; i < toSet.Length; ++i)
            {
                ToolStripMenuItem newItem = new ToolStripMenuItem(toSet[i]);
                newItem.Click += new System.EventHandler(this.insertMergeField);
                insertMergeFieldToolStripMenuItem.DropDownItems.Add(newItem);
            }
        }
        private void insertMergeField(object sender, System.EventArgs e)
        {
            bodyBox.SelectedRtf = "{\\rtf1" + "[[" + sender + "]]" + "}}";
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
    }
}
