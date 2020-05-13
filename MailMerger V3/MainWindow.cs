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
        List<string> toDelete = new List<string>();
        OpenFileDialog getAttachment = new OpenFileDialog();
        OpenFileDialog getImage = new OpenFileDialog();

        //Settings
        string rtfSignature;
        //string htmlSignature;
        string htmlEmailBody;
        string[] referenceAliases;
        Font defaultFont;
        //bool signatureContainsSignoff = false;

        //User info
        string emailAddress;
        string firstName;
        string jobTitle;
        string department;
        string phoneNumber;

        public main()
        {
            InitializeComponent();
            logBox.SelectAll();
            logBox.SelectionAlignment = HorizontalAlignment.Center;
            //Get the email address
            inboxes.Items.Add(getEmail());
            //EWS setup
            service.UseDefaultCredentials = true;
            startAutoDiscover = new Thread(autoDiscoverThread);
            startAutoDiscover.IsBackground = true;
            startAutoDiscover.Start();
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
                setupForm.Close();
            }
        }

        private void importButton_Click(object sender, EventArgs e)
        {
            openBook.Multiselect = false;
            openBook.Filter = "Excel Files | *.csv; *.xls*";
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
                var previewEmail = createEmail(allData[currentRecipient + 1], currentRecipient + 1);
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
            } else if (String.IsNullOrEmpty(subjectBox.Text.Trim()) || subjectBox.Text.Trim().Equals("[Replace with subject]"))
            {
                writeLog("Please change the email subject before attempting to send.");
            } else if (String.IsNullOrEmpty(bodyBox.Text.Trim()) || bodyBox.Text.Trim().Equals("[Replace with body of email]"))
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

        private void settingsButton_Click(object sender, EventArgs e)
        {
            Setup setupForm = new Setup();
            setupForm.ShowDialog();
            importSettings();
            setupForm.Close();
        }

        private void addButton_Click(object sender, EventArgs e)
        {
            getAttachment.Multiselect = false;
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
            logBox.Invoke((Action)delegate

            {

                logBox.Text = logBox.Text + toWrite + Environment.NewLine + Environment.NewLine;
                logBox.SelectAll();
                logBox.SelectionAlignment = HorizontalAlignment.Center;
                logBox.Select(logBox.Text.Length - 1, 1);
                logBox.ScrollToCaret();
            });
        }
        private static  void writeErrors(string toWrite)
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
        private void ResetForImport()
        {
            insertMergeFieldToolStripMenuItem.DropDownItems.Clear();
            importSuccess = false;
            forenameColumn = -1;
            emailColumn = -1;
            referenceColumn = -1;
            dearBox.Text = "";
            refBox.Text = "";
            emailBox.Text = "";
        }
        private void importWorkbook(string filepath)
        {
            ResetForImport();
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
                sheetSelect.Close();
            } else
            {
                importSheet = importBook.Worksheets[1];
            }
            usedRange = importSheet.UsedRange;
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
            FindColumns();
        }
        private void FindColumns()
        {
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
                    if ((!(String.IsNullOrEmpty(referenceAliases[0]))) && UsefulTools.findMatch(allData[0][i], referenceAliases))
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
                }
                else
                {
                    writeLog("Import successful!");
                }
                importSuccess = true;
                UsefulTools.setMergeFields(allData[0], insertMergeFieldToolStripMenuItem, bodyBox);
            }
            else if (forenameColumn == -1)
            {
                writeLog("Import failed - forename column missing.");
            }
            else if (emailColumn == -1)
            {
                writeLog("Import failed - email column missing.");
            }
        }

        private void main_FormClosing(object sender, FormClosingEventArgs e)
        {
            try
            {
                importBook.Close();
            } catch (System.NullReferenceException)
            {
                //Excel workbook never opened - do nothing!
            } catch (System.Runtime.InteropServices.COMException)
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
            if (toDelete != null)
            {
                //Delete created resources
                foreach (string eachString in toDelete)
                {
                    if (File.Exists(eachString))
                    {
                        File.Delete(eachString);
                    }
                }
            }
            excel.Quit();
            //outlook.Quit();
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
            while (initialiseSuccess == false)
            {
                Thread.Sleep(500);
            }
            try
            {
                service.AutodiscoverUrl(getEmail(), autoDiscoverOverload);
                autoDiscoverFinished = true;
                writeLog("Finished configuring your Exchange profile!");
            } catch (Exception)
            {
                string password = "";
                while (password == null || String.IsNullOrEmpty(password))
                {
                    this.Invoke((Action)(() => { 
                    
                    ModalPrompt passwordPrompt = new ModalPrompt("Password", "Enter your password", true);
                    passwordPrompt.ShowDialog();
                    password = passwordPrompt.Result;
                    if (password == null)
                    {
                        Application.Exit();
                        Environment.Exit(0);
                    }
                    }));
                }
                service.UseDefaultCredentials = false;
                service.Credentials = new WebCredentials(emailAddress, password);
                try
                {
                    service.AutodiscoverUrl(getEmail(), autoDiscoverOverload);
                    autoDiscoverFinished = true;
                    writeLog("Finished configuring your Exchange profile!");
                } catch (Exception e)
                {
                    MessageBox.Show("Error!: " + e.Message + ". Unable to login to Exchange profile!");
                    Application.Exit();
                    Environment.Exit(0);
                }
            }
            
        }

        private bool autoDiscoverOverload(string redirectionUrl)
        {
            // The default for the validation callback is to reject the URL.
            bool result = false;
            Uri redirectionUri = new Uri(redirectionUrl);
            // Validate the contents of the redirection URL. In this simple validation
            // callback, the redirection URL is considered valid if it is using HTTPS
            // to encrypt the authentication credentials. 
            if (redirectionUri.Scheme == "https")
            {
                result = true;
            }
            return result;
        }
        private EmailMessage createEmail(string[] recipientData, int recipientNumber)
        {
            EmailMessage email = new EmailMessage(service);
            int imageOccurences;
            int lastImgOccurence = 0;
            string thisBody = "<span style=\"font-family:" + defaultFont.OriginalFontName + "; font-size:" + defaultFont.Size.ToString() + "pt;\">Dear " + recipientData[forenameColumn] + ",</span>" + "<br/><br/>";
            thisBody += replaceMergeFields(htmlEmailBody, recipientNumber);
            imageOccurences = UsefulTools.CountStringOccurrences(thisBody, "<img");
            //Attach all images and embed them into html body
            for (int i = 0; i < imageOccurences; ++i)
            {
                lastImgOccurence = thisBody.IndexOf("<img", lastImgOccurence + 1);
                string imgSRC = UsefulTools.grabImgSRC(lastImgOccurence, thisBody);
                string fileLocation = imgSRC.Substring(1);
                //htmlEmailBody = htmlEmailBody.Replace(imgSRC, "CID:" + fileLocation.Substring(1));
                email.Attachments.AddFileAttachment(fileLocation);
                email.Attachments[i].IsInline = true;
                email.Attachments[i].ContentId = fileLocation.Substring(1);
                toDelete.Add(fileLocation);
            }
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
            string link;
            ModalPrompt linkPrompt = new ModalPrompt("Hyperlink", "Enter link for hyperlink", false);
            linkPrompt.ShowDialog();
            link = linkPrompt.Result;
            if(link != null & String.IsNullOrEmpty(link))
            {
                return;
            }
            UsefulTools.replaceHighlightedRtf("\\cf3\\ul " + bodyBox.SelectedText + " <" + link + ">\\cf7\\ulnone", bodyBox);
        }

        private void pasteToolStripMenuItem_Click(object sender, EventArgs e)
        {
            bodyBox.SelectedRtf = Clipboard.GetData(DataFormats.Rtf).ToString();
        }
        private void insertImageToolStripMenuItem_Click(object sender, EventArgs e)
        {
            getImage.Multiselect = false;
            getImage.Filter = "Image Files | *.png; *.jpeg; *.bmp; *.gif";
            Image image;
            if (getImage.ShowDialog() == DialogResult.OK && File.Exists(getImage.FileName))
            {
                try
                {
                    image = Image.FromFile(getImage.FileName);
                    Clipboard.SetImage(image);
                    bodyBox.Paste();
                } catch (OutOfMemoryException)
                {
                    writeLog("Please select a valid image.");
                }
            }
            else
            {
                writeLog("Please select a valid image.");
            }
        }

        private void emailPrep()
        {
            string parseRtf = "";
            HtmlResult parsedHTML;
            UsefulTools.trimRtf(bodyBox);
            bodyBox.SelectAll();
            bodyBox.Font = defaultFont;
            if (bodyBox.Text.ToLower().Trim().EndsWith("kind regards") || bodyBox.Text.ToLower().Trim().EndsWith("best wishes") || bodyBox.Text.ToLower().Trim().EndsWith("all the best") || bodyBox.Text.ToLower().Trim().EndsWith("best regards") || bodyBox.Text.ToLower().Trim().EndsWith("yours sincerely"))
            {
                bodyBox.AppendText(Environment.NewLine);
                parseRtf = bodyBox.Rtf + rtfSignature;
            } else
            {
                bodyBox.AppendText(Environment.NewLine + Environment.NewLine + "Kind regards" + Environment.NewLine);
                parseRtf = bodyBox.Rtf + rtfSignature;
            }
            parsedHTML = RtfToHtmlConverter.RtfToHtml(parseRtf);
            parsedHTML.WriteToFile("emailHTML.html");
            htmlEmailBody = parsedHTML.Html;
            htmlEmailBody = replaceSignatureFields(htmlEmailBody);
            string sendingInbox = inboxes.SelectedItem.ToString().Trim();
            Mailbox sentBox = new Mailbox(sendingInbox);
            sentBoxSentItems = new FolderId(WellKnownFolderName.SentItems, sentBox);
            sentBoxDrafts = new FolderId(WellKnownFolderName.Drafts, sentBox);
            sendFrom = new EmailAddress(sendingInbox);
        }
        private void sendAll(object sender, DoWorkEventArgs e)
        {
            for (int i = 1; i < allData.Length; ++i)
            {
                if (String.IsNullOrEmpty(allData[i][forenameColumn].Trim()) && String.IsNullOrEmpty(allData[i][emailColumn].Trim()))
                {
                    //Do nothing
                }
                else if (String.IsNullOrEmpty(allData[i][forenameColumn].Trim()))
                {
                    errors += ("Error for recipient on row " + i + " (" + allData[i][emailColumn] + ") - forename is missing." + Environment.NewLine);
                }
                else if (String.IsNullOrEmpty(allData[i][emailColumn].Trim()))
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
            string defaultFontString = "";
            SettingsImporter.importSettings(ref rtfSignature, ref inboxes, ref stringAliases, ref defaultFontString);
            defaultFont = new Font(UsefulTools.getFontValue(defaultFontString, "Name"), int.Parse(UsefulTools.getFontValue(defaultFontString, "Size")));
            if (stringAliases.Trim().Equals("Replace with comma separated list of customer reference aliases") || (String.IsNullOrEmpty(stringAliases.Trim())))
            {
                referenceAliases = new string[] { "" };
            } else if (stringAliases.IndexOf(',') != -1)
            {
                referenceAliases = stringAliases.Split(',').Select(p => p.Trim()).ToArray();
            } else
            {
                referenceAliases = new string[] {stringAliases};
            }
            bodyBox.SelectAll();
            bodyBox.SelectionFont = defaultFont;
            bodyBox.Select(0, 0);
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
                toReplace = toReplace.Replace("[[Email]]", inboxes.SelectedItem.ToString());
            }
            if (toReplace.Contains("[[Job Title]]"))
            {
                toReplace = toReplace.Replace("[[Job Title]]", getJob());
            }
            if (toReplace.Contains("[[Department]]"))
            {
                toReplace = toReplace.Replace("[[Department]]", getDepartment());
            }
            if (toReplace.Contains("[[Phone Number]]"))
            {
                toReplace = toReplace.Replace("[[Phone Number]]", getPhone());
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
            
            if (String.IsNullOrEmpty(e.Result.ToString().Trim())) {
                writeLog("Emails sent without errors!");
            } else
            {
                writeLog("Emails sent with errors:" + Environment.NewLine + Environment.NewLine + e.Result.ToString());
                writeErrors(e.Result.ToString());
            }
            errors = "";
        }

        private string getName()
        {
            if (String.IsNullOrEmpty(firstName))
            {
                while (!autoDiscoverFinished)
                {
                    Thread.Sleep(100);
                }
                try
                {
                    firstName = service.ResolveName(getEmail().Substring(0, getEmail().IndexOf('@')), ResolveNameSearchLocation.DirectoryOnly, true)[0].Contact.DisplayName;
                }
                catch (Exception)
                {
                    while (firstName == null || String.IsNullOrEmpty(firstName.Trim()))
                    {
                        ModalPrompt firstNamePrompt = new ModalPrompt("Name Required", "Name not found! Please enter your name.", false);
                        firstNamePrompt.ShowDialog();
                        firstName = firstNamePrompt.Result;
                    }
                }
            }
            return firstName;
        }

        private  string getEmail()
        {
            if (String.IsNullOrEmpty(emailAddress))
            {
                try
                {
                    emailAddress = System.DirectoryServices.AccountManagement.UserPrincipal.Current.EmailAddress;
                    while (String.IsNullOrEmpty(emailAddress))
                    {
                        ModalPrompt emailPrompt = new ModalPrompt("Email Address", "Please provide your email address", false);
                        emailPrompt.ShowDialog();
                        emailAddress = emailPrompt.Result;
                        if (emailAddress == null)
                        {
                            Application.Exit();
                            Environment.Exit(0);
                        }
                    }

                }
                catch (Exception e)
                {
                    MessageBox.Show("Error: " + e.Message + ".");
                }
            }
            return emailAddress;
        }

        private string getJob()
        {
            if (String.IsNullOrEmpty(jobTitle))
            {
                while (!autoDiscoverFinished)
                {
                    Thread.Sleep(100);
                }
                try
                {
                    jobTitle = service.ResolveName(getEmail().Substring(0, getEmail().IndexOf('@')), ResolveNameSearchLocation.DirectoryOnly, true)[0].Contact.JobTitle;
                }
                catch (Exception)
                {
                    while (jobTitle == null || String.IsNullOrEmpty(jobTitle.Trim()))
                    {
                        ModalPrompt jobPrompt = new ModalPrompt("Job Title Required", "Job title not found! Please enter your job title.", false);
                        jobPrompt.ShowDialog();
                        jobTitle = jobPrompt.Result;
                    }
                }
            }
            return jobTitle;
        }
        private string getPhone()
        {
            if (String.IsNullOrEmpty(phoneNumber))
            {
                while (!autoDiscoverFinished)
                {
                    Thread.Sleep(100);
                }
                try
                {
                    phoneNumber = service.ResolveName(getEmail().Substring(0, getEmail().IndexOf('@')), ResolveNameSearchLocation.DirectoryOnly, true)[0].Contact.PhoneNumbers[PhoneNumberKey.BusinessPhone];
                }
                catch (Exception)
                {
                    while (phoneNumber == null || String.IsNullOrEmpty(phoneNumber.Trim()))
                    {
                        ModalPrompt phonePrompt = new ModalPrompt("Phone number Required", "Phone number not found! Please enter your phone number.", false);
                        phonePrompt.ShowDialog();
                        phoneNumber = phonePrompt.Result;
                    }
                }
            }
            return phoneNumber;
        }
        private string getDepartment()
        {
            if (String.IsNullOrEmpty(department))
            {
                while (!autoDiscoverFinished)
                {
                    Thread.Sleep(100);
                }
                try
                {
                    department = service.ResolveName(getEmail().Substring(0, getEmail().IndexOf('@')), ResolveNameSearchLocation.DirectoryOnly, true)[0].Contact.Department;
                }
                catch (Exception)
                {
                    while (department == null || String.IsNullOrEmpty(department.Trim()))
                    {
                        ModalPrompt departmentPrompt = new ModalPrompt("Department Required", "Department not found! Please enter your department.", false);
                        departmentPrompt.ShowDialog();
                        department = departmentPrompt.Result;
                    }
                }
            }
            return department;
        }
    }
}
