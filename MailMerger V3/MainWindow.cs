using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Exchange.WebServices.Data;
using System.Threading;
using Rtf2Html;
using System.IO;
using System.Runtime.InteropServices;

namespace MailMerger_V3
{
    public partial class main : Form
    {
        //Main application setup
        private readonly Excel.Application _excel = new Excel.Application();
        private readonly OpenFileDialog _openBook = new OpenFileDialog();
        private int _forenameColumn = -1, _referenceColumn = -1, _emailColumn = -1, _currentRecipient = 0;
        private Excel.Workbook _importBook;
        private Excel.Worksheet _importSheet;
        private Excel.Workbooks _workbooks;
        private Excel.Range _usedRange;
        private string[][] AllData { get; set; }
        private bool _importSuccess = false;
        private bool InitialiseSuccess { get; set; } = false;
        private string _errors = "";

        //EWS stuff
        private readonly ExchangeService _service = new ExchangeService();
        private bool _autoDiscoverFinished = false;
        private EmailAddress _sendForm;
        private FolderId _sentBoxSentItems;
        private FolderId _sentBoxDrafts;

        //Related to files and attachments
        private readonly List<string> _toDelete = new List<string>();
        private readonly OpenFileDialog _getAttachment = new OpenFileDialog();
        private readonly OpenFileDialog _getImage = new OpenFileDialog();

        //Settings
        private string _rtfSignature;
        //string htmlSignature;
        private string _htmlEmailBody;
        private string[] _referenceAliases;
        private Font _defaultFont;
        //bool signatureContainsSignoff = false;

        //User info
        private readonly string _userEmailAddress;
        private readonly string _userPassword;
        private string _firstName;
        private string _jobTitle;
        private string _department;
        private string _phoneNumber;

        //Other
        string bodyFontFamily;
        string bodyFontSize;

        public main()
        {
            InitializeComponent();
            //Get user credentials
            UserLogin loginForm = new UserLogin();
            loginForm.ShowDialog();
            //Close if not provided
            if (loginForm.DialogResult == DialogResult.OK)
            {
                _userEmailAddress = loginForm.UserEmail;
                _userPassword = loginForm.UserPassword;
            }
            else
            {
                Application.Exit();
                Environment.Exit(0);
            }

            logBox.SelectAll();
            logBox.SelectionAlignment = HorizontalAlignment.Center;
            //Get the email address
            inboxes.Items.Add(_userEmailAddress);
            //EWS setup
            _service.UseDefaultCredentials = false;
            _service.Credentials = new WebCredentials(_userEmailAddress, _userPassword);
            Thread startAutoDiscover = new Thread(AutoDiscoverThread)
            {
                IsBackground = true
            };
            startAutoDiscover.Start();
            if (File.Exists("Settings.txt"))
            {
                ImportSettings();
                InitialiseSuccess = true;
            }
            else
            {
                //Start setup
                Setup setupForm = new Setup();
                setupForm.ShowDialog();
                if (setupForm.DialogResult == DialogResult.OK)
                {
                    ImportSettings();
                    InitialiseSuccess = true;
                }
                setupForm.Close();
            }
        }

        private void ImportButton_Click(object sender, EventArgs e)
        {
            _openBook.Multiselect = false;
            _openBook.Filter = @"Excel Files | *.csv; *.xls*";
            if (_openBook.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    ImportWorkbook(_openBook.FileName);
                    if (_importSuccess)
                    {
                        DisplayRecipient(0);
                    }
                    //This line causes a crash on Windows 7 when using formulas in the worksheet
                    CleanupExcel();
                } catch (Microsoft.CSharp.RuntimeBinder.RuntimeBinderException)
                {
                    WriteLog("Selected file is not a valid excel file.");
                }
            } else
            {
                WriteLog("Please select an excel file.");
            }
        }

        private void CleanupExcel()
        {
            try
            {
                Marshal.ReleaseComObject(_importSheet);
            }
            catch (Exception)
            {
                //Unable to close sheet, but not the biggest issue in the world!
            }
            try
            {
                Marshal.ReleaseComObject(_importBook);
            }
            catch (Exception)
            {
                //Unable to close workbook, but not the biggest issue in the world!
            }
            try
            {
                Marshal.ReleaseComObject(_workbooks);
            }
            catch (Exception)
            {
                //Unable to close Excel, but not the biggest issue in the world!
            }
        }
        private void previewButton_Click(object sender, EventArgs e)
        {
            if (!_autoDiscoverFinished)
            {
                WriteLog("Please wait for your exchange profile to configure before previewing.");
                return;
            }
            if (_importSuccess)
            {
                EmailPrep();
                EmailMessage previewEmail = CreateEmail(AllData[_currentRecipient + 1], _currentRecipient + 1);
                previewEmail.Save(WellKnownFolderName.Drafts);
                PropertySet psPropSet = new PropertySet(ItemSchema.MimeContent);
                EmailMessage emEmailMessage = EmailMessage.Bind(_service, previewEmail.Id, psPropSet);
                FileStream fsFileStream = new FileStream("preview.eml", FileMode.Create);
                fsFileStream.Write(emEmailMessage.MimeContent.Content, 0, emEmailMessage.MimeContent.Content.Length);
                fsFileStream.Close();
                previewEmail.Delete(DeleteMode.HardDelete);
                System.Diagnostics.Process.Start("preview.eml");
            } else
            {
                WriteLog("Please import before previewing.");
            }
            
        }

        private void sendButton_Click(object sender, EventArgs e)
        {
            if (_importSuccess == false)
            {
                WriteLog("Please import data before attempting to send emails.");
            } else if (_autoDiscoverFinished == false)
            {
                WriteLog("System is still configuring your exchange profile - please wait for a couple of minutes and try again.");
            } else if (String.IsNullOrEmpty(subjectBox.Text.Trim()) || subjectBox.Text.Trim().Equals("[Replace with subject]"))
            {
                WriteLog("Please change the email subject before attempting to send.");
            } else if (String.IsNullOrEmpty(bodyBox.Text.Trim()) || bodyBox.Text.Trim().Equals("[Replace with body of email]"))
            {
                WriteLog("Please change the email body before attempting to send.");
            } else
            {
                if (MessageBox.Show(@"Are you sure you want to send?", @"Confirm send", MessageBoxButtons.YesNo) == DialogResult.Yes)
                {
                    EnableForm(false);
                    EmailPrep();
                    //Update on progress
                    BackgroundWorker backgroundSend = new BackgroundWorker();
                    backgroundSend.DoWork += SendAll;
                    backgroundSend.RunWorkerCompleted += SendFinished;
                    backgroundSend.RunWorkerAsync();

                }
            }
        }

        private void settingsButton_Click(object sender, EventArgs e)
        {
            Setup setupForm = new Setup();
            setupForm.ShowDialog();
            ImportSettings();
            setupForm.Close();
        }

        private void addButton_Click(object sender, EventArgs e)
        {
            _getAttachment.Multiselect = false;
            if (_getAttachment.ShowDialog() == DialogResult.OK && File.Exists(_getAttachment.FileName))
            {

                attachments.Items.Add(new Attachment(_getAttachment.FileName));
                attachments.SelectedItem = attachments.Items[attachments.Items.Count - 1];
            } else
            {
                WriteLog("Please select a valid file to attach.");
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

        private void NextButton_Click(object sender, EventArgs e)
        {
            if (_importSuccess && _currentRecipient < AllData.Length - 2)
            {
                DisplayRecipient(_currentRecipient + 1);
            }
        }

        private void PreviousButton_Click(object sender, EventArgs e)
        {
            if (_importSuccess && _currentRecipient > 0)
            {
                DisplayRecipient(_currentRecipient - 1);
            }
        }

        private void ClearButton_Click(object sender, EventArgs e)
        {
            bodyBox.Text = "";
            subjectBox.Text = "";
        }
        private void WriteLog(string toWrite)
        {
            logBox.Text = logBox.Text + toWrite + Environment.NewLine + Environment.NewLine;
            logBox.SelectAll();
            logBox.SelectionAlignment = HorizontalAlignment.Center;
            logBox.Select(logBox.Text.Length - 1, 1);
            logBox.ScrollToCaret();
        }
        private static  void WriteErrors(string toWrite)
        {
            if (File.Exists("Errors.txt"))
            {
                File.AppendAllText("Errors.txt", @"Error occured at " + DateTime.Now + @":" + Environment.NewLine + toWrite + Environment.NewLine);
            } else
            {
                File.Create("Errors.txt").Close();
                File.AppendAllText("Errors.txt", @"Error occured at " + DateTime.Now + @":" + Environment.NewLine + toWrite + Environment.NewLine);
            }
        }
        private void ResetForImport()
        {
            insertMergeFieldToolStripMenuItem.DropDownItems.Clear();
            _importSuccess = false;
            _forenameColumn = -1;
            _emailColumn = -1;
            _referenceColumn = -1;
            dearBox.Text = "";
            refBox.Text = "";
            emailBox.Text = "";
        }
        private void ImportWorkbook(string filepath)
        {
            try
            {
                ResetForImport();
                _workbooks = _excel.Workbooks;
                _importBook = _workbooks.Open(filepath);
                if (_importBook.Worksheets.Count > 1)
                {
                    string[] possibleSheets = new string[_importBook.Worksheets.Count];
                    for (int i = 1; i <= _importBook.Worksheets.Count; ++i)
                    {
                        possibleSheets[i - 1] = _importBook.Worksheets[i].Name;
                    }

                    SheetSelect sheetSelect = new SheetSelect(possibleSheets);
                    while (sheetSelect.DialogResult != DialogResult.OK)
                    {
                        sheetSelect.ShowDialog();
                    }

                    _importSheet = _importBook.Worksheets[sheetSelect.selectedSheet];
                    sheetSelect.Close();
                }
                else
                {
                    _importSheet = _importBook.Worksheets[1];
                }

                _usedRange = _importSheet.UsedRange;
                AllData = new string[_usedRange.Rows.Count][];
                //Get all sheet data
                for (int row = 0; row < _usedRange.Rows.Count; ++row)
                {
                    AllData[row] = new string[_usedRange.Columns.Count];
                    for (int column = 0; column < _usedRange.Columns.Count; ++column)
                    {
                        if (_importSheet.Cells[row + 1, column + 1].Value == null)
                        {
                            AllData[row][column] = "";
                        }
                        else
                        {
                            AllData[row][column] = _importSheet.Cells[row + 1, column + 1].Value.ToString().Trim();
                        }
                    }
                }

                FindColumns();
            }
            catch (Exception e)
            {
                MessageBox.Show(@"Unknown import error: " + e.Message + @" at: " + e.StackTrace);
                Application.Exit();
                Environment.Exit(0);
            }
        }
        private void FindColumns()
        {
            //Get column headers
            for (int i = 0; i < AllData[0].Length; ++i)
            {
                if (_forenameColumn == -1)
                {
                    if (UsefulTools.FindMatch(AllData[0][i], new[] { "forename", "firstname", "first name", "fore name" }))
                    {
                        _forenameColumn = i;
                        WriteLog("Forename is column " + (_forenameColumn + 1));
                    }
                }
                if (_referenceColumn == -1)
                {
                    if ((!(String.IsNullOrEmpty(_referenceAliases[0]))) && UsefulTools.FindMatch(AllData[0][i], _referenceAliases))
                    {
                        _referenceColumn = i;
                        WriteLog(_referenceAliases[0] + " is column " + (_referenceColumn + 1));
                    }
                }
                if (_emailColumn == -1)
                {
                    if (UsefulTools.FindMatch(AllData[0][i], new[] { "email", " e-mail" }))
                    {
                        _emailColumn = i;
                        WriteLog("Email is column " + (_emailColumn + 1));
                    }
                }
            }
            if (_forenameColumn != -1 && _emailColumn != -1)
            {
                WriteLog("Number of recipients: " + (AllData.Length - 1));
                if (_referenceColumn == -1)
                {
                    WriteLog("Import was successful - but reference column was not found.");
                }
                else
                {
                    WriteLog("Import successful!");
                }
                _importSuccess = true;
                UsefulTools.SetMergeFields(AllData[0], insertMergeFieldToolStripMenuItem, bodyBox);
            }
            else if (_forenameColumn == -1)
            {
                WriteLog("Import failed - forename column missing.");
            }
            else if (_emailColumn == -1)
            {
                WriteLog("Import failed - email column missing.");
            }
        }

        private void main_FormClosing(object sender, FormClosingEventArgs e)
        {
            CleanupExcel();
            if (File.Exists("preview.eml"))
            {
                File.Delete("preview.eml");
            }
            if (File.Exists("emailHTML.html"))
            {
                File.Delete("emailHTML.html");
            }
            if (_toDelete != null)
            {
                //Delete created resources
                foreach (string eachString in _toDelete)
                {
                    if (File.Exists(eachString))
                    {
                        File.Delete(eachString);
                    }
                }
            }
            _excel.Quit();
        }
        private void DisplayRecipient(int toDisplay)
        {
            _currentRecipient = toDisplay;
            dearBox.Text = AllData[toDisplay + 1][_forenameColumn];
            if (_referenceColumn != - 1)
            {
                refBox.Text = AllData[toDisplay + 1][_referenceColumn];
            }
            emailBox.Text = AllData[toDisplay + 1][_emailColumn];
        }
        private void AutoDiscoverThread()
        {
            try
            {
                while (InitialiseSuccess == false)
                {
                    Thread.Sleep(500);
                }
                _service.AutodiscoverUrl(_userEmailAddress, autoDiscoverOverload);
                _autoDiscoverFinished = true;
                MessageBox.Show(@"Finished configuring your Exchange profile.");
            }
            catch (Exception e)
            {
                MessageBox.Show(@"Error: " + e.Message + @". Unable to login to Exchange profile!");
                Application.Exit();
                Environment.Exit(0);
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
        private EmailMessage CreateEmail(string[] recipientData, int recipientNumber)
        {
            EmailMessage email = new EmailMessage(_service);
            int lastImgOccurence = 0;
            //string thisBody = "<span style=\"font-family:" + _defaultFont.OriginalFontName + "; font-size:" + _defaultFont.Size + "pt;\">Dear " + recipientData[_forenameColumn] + ",</span>" + "<br/><br/>";
            string thisBody = "<span style=\"font-family:" + bodyFontFamily + "; font-size:" + bodyFontSize + ";\">Dear " + recipientData[_forenameColumn] + ",</span>" + "<br/><br/>";
            thisBody += ReplaceMergeFields(_htmlEmailBody, recipientNumber);
            int imageOccurences = UsefulTools.CountStringOccurrences(thisBody, "<img");
            //Attach all images and embed them into html body
            for (int i = 0; i < imageOccurences; ++i)
            {
                lastImgOccurence = thisBody.IndexOf("<img", lastImgOccurence + 1);
                string imgLocation = UsefulTools.GrabImgSrc(lastImgOccurence, thisBody);
                string imgName = imgLocation.Substring(2);
                thisBody = thisBody.Replace(imgLocation, "cid:" + imgName);
                email.Attachments.AddFileAttachment(imgName, imgLocation);
                email.Attachments[i].IsInline = true;
                email.Attachments[i].ContentId = imgName;
                _toDelete.Add(imgLocation);
            }
            email.Body = thisBody;
            email.Subject = subjectBox.Text;
            if (_referenceColumn != -1)
            {
                email.Subject += " - " + _referenceAliases[0] + ": " + recipientData[_referenceColumn];
            }
            email.ToRecipients.Add(recipientData[_emailColumn]);
            //email.Sender = _sendForm;
            email.From = _sendForm;
            //Attach attachments
            foreach (Attachment eachAttachment in attachments.Items)
            {
                if (File.Exists(eachAttachment.FilePath))
                {
                    email.Attachments.AddFileAttachment(eachAttachment.FilePath);
                } else
                {
                    _errors += ("Error for recipient on row " + recipientNumber + " (" + AllData[recipientNumber][_forenameColumn] + ") - attachment \"" + eachAttachment.FileName + "\" cannot be found." + Environment.NewLine);
                }
            }
            return email;
        }

        private void main_Load(object sender, EventArgs e)
        {
            if(!(InitialiseSuccess))
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
            ModalPrompt linkPrompt = new ModalPrompt("Hyperlink", "Enter link for hyperlink");
            linkPrompt.ShowDialog();
            link = linkPrompt.Result;
            if(link != null & String.IsNullOrEmpty(link))
            {
                return;
            }
            UsefulTools.ReplaceHighlightedRtf("\\cf3\\ul " + bodyBox.SelectedText + " <" + link + ">\\cf7\\ulnone", bodyBox);
        }

        private void pasteToolStripMenuItem_Click(object sender, EventArgs e)
        {
            bodyBox.SelectedRtf = Clipboard.GetData(DataFormats.Rtf).ToString();
        }
        private void insertImageToolStripMenuItem_Click(object sender, EventArgs e)
        {
            _getImage.Multiselect = false;
            _getImage.Filter = @"Image Files | *.png; *.jpeg; *.bmp; *.gif";
            if (_getImage.ShowDialog() == DialogResult.OK && File.Exists(_getImage.FileName))
            {
                try
                {
                    Image image = Image.FromFile(_getImage.FileName);
                    Clipboard.SetImage(image);
                    bodyBox.Paste();
                } catch (OutOfMemoryException)
                {
                    WriteLog("Please select a valid image.");
                }
            }
            else
            {
                WriteLog("Please select a valid image.");
            }
        }

        private void EmailPrep()
        {
            string parseRtf;
            UsefulTools.TrimRtf(bodyBox);
            //Change the font
            bodyBox.SelectAll();
            //This overwrites bold etc...
            //bodyBox.Font = _defaultFont; 
            //Grab the font being used by the body
            bodyFontFamily = UsefulTools.GetFontValue(RtfToHtmlConverter.RtfToHtml(bodyBox.Rtf).Html, "font-family");
            bodyFontSize = UsefulTools.GetFontValue(RtfToHtmlConverter.RtfToHtml(bodyBox.Rtf).Html, "font-size");
            //Change the signature to use this font
            using (RichTextBox tempRichbox = new RichTextBox())
            {
                tempRichbox.Rtf = _rtfSignature;
                if (bodyBox.Text.ToLower().Trim().EndsWith("kind regards") || bodyBox.Text.ToLower().Trim().EndsWith("best wishes") || bodyBox.Text.ToLower().Trim().EndsWith("all the best") || bodyBox.Text.ToLower().Trim().EndsWith("best regards") || bodyBox.Text.ToLower().Trim().EndsWith("yours sincerely"))
                {
                    tempRichbox.SelectionStart = 0;
                    tempRichbox.SelectionLength = 0;
                    tempRichbox.SelectedText = (Environment.NewLine);
                }
                else
                {
                    tempRichbox.SelectionStart = 0;
                    tempRichbox.SelectionLength = 0;
                    tempRichbox.SelectedText =  (Environment.NewLine + "Kind regards" + Environment.NewLine + Environment.NewLine);
                }
                tempRichbox.SelectAll();
                tempRichbox.Font = new Font(bodyFontFamily, float.Parse(bodyFontSize.Replace("pt", "")));
                parseRtf = bodyBox.Rtf + tempRichbox.Rtf;
            }
            HtmlResult parsedHtml = RtfToHtmlConverter.RtfToHtml(parseRtf);
            parsedHtml.WriteToFile("emailHTML.html");
            _htmlEmailBody = parsedHtml.Html;
            _htmlEmailBody = ReplaceSignatureFields(_htmlEmailBody);
            string sendingInbox = inboxes.SelectedItem.ToString().Trim();
            Mailbox sentBox = new Mailbox(sendingInbox);
            _sentBoxSentItems = new FolderId(WellKnownFolderName.SentItems, sentBox);
            _sentBoxDrafts = new FolderId(WellKnownFolderName.Drafts, sentBox);
            _sendForm = new EmailAddress(sendingInbox);
        }
        private void SendAll(object sender, DoWorkEventArgs e)
        {
            for (int i = 1; i < AllData.Length; ++i)
            {
                if (String.IsNullOrEmpty(AllData[i][_forenameColumn].Trim()) && String.IsNullOrEmpty(AllData[i][_emailColumn].Trim()))
                {
                    //Do nothing
                }
                else if (String.IsNullOrEmpty(AllData[i][_forenameColumn].Trim()))
                {
                    _errors += ("Error for recipient on row " + i + " (" + AllData[i][_emailColumn] + ") - forename is missing." + Environment.NewLine);
                }
                else if (String.IsNullOrEmpty(AllData[i][_emailColumn].Trim()))
                {
                    _errors += ("Error for recipient on row " + i + " (" + AllData[i][_forenameColumn] + ") - email address is missing." + Environment.NewLine);
                }
                else if (!(UsefulTools.IsValidEmail(AllData[i][_emailColumn])))
                {
                    _errors += ("Error for recipient on row " + i + " (" + AllData[i][_forenameColumn] + ") - email address is invalid." + Environment.NewLine);
                }  else
                {
                    SendEmail(AllData[i], i);
                }
            }
            e.Result = _errors;
        }
        private void SendEmail(String[] recipientData, int recipientNumber)
        {
            EmailMessage finalEmail = CreateEmail(recipientData, recipientNumber);
            try
            {
                finalEmail.Save(_sentBoxDrafts);
                finalEmail.SendAndSaveCopy(_sentBoxSentItems);
            } catch (Exception)
            {
                try
                {
                    finalEmail.Save();
                    finalEmail.SendAndSaveCopy();
                }
                catch (Exception e)
                {
                    MessageBox.Show(@"Unrecoverable error: " + e.Message + "At:" + e.StackTrace);
                    Application.Exit();
                    Environment.Exit(0);
                }
            }
        }
        private void ImportSettings()
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
            SettingsImporter.ImportSettings(ref _rtfSignature, ref inboxes, ref stringAliases, ref defaultFontString);
            _defaultFont = new Font(UsefulTools.GetFontValue(defaultFontString, "Name"), int.Parse(UsefulTools.GetFontValue(defaultFontString, "Size")));
            if (stringAliases.Trim().Equals("Replace with comma separated list of customer reference aliases") || (String.IsNullOrEmpty(stringAliases.Trim())))
            {
                _referenceAliases = new[] { "" };
            } else if (stringAliases.IndexOf(',') != -1)
            {
                _referenceAliases = stringAliases.Split(',').Select(p => p.Trim()).ToArray();
            } else
            {
                _referenceAliases = new[] {stringAliases};
            }
            bodyBox.SelectAll();
            bodyBox.SelectionFont = _defaultFont;
            bodyBox.Select(0, 0);
        }
        private string ReplaceMergeFields(string toReplace, int recipientNumber)
        {
            for (int i = 0; i < AllData[0].Length; ++i)
            {
                //for each header
                if (toReplace.Contains("[[" + AllData[0][i] + "]]"))
                {
                    toReplace = toReplace.Replace("[[" + AllData[0][i] + "]]", AllData[recipientNumber][i]);
                }
            }
            return toReplace;
        }
        private string ReplaceSignatureFields(string toReplace)
        {
            if (toReplace.Contains("[[Name]]"))
            {
                toReplace = toReplace.Replace("[[Name]]", GetName());
            }
            if (toReplace.Contains("[[Email]]"))
            {
                toReplace = toReplace.Replace("[[Email]]", inboxes.SelectedItem.ToString());
            }
            if (toReplace.Contains("[[Job Title]]"))
            {
                toReplace = toReplace.Replace("[[Job Title]]", GetJob());
            }
            if (toReplace.Contains("[[Department]]"))
            {
                toReplace = toReplace.Replace("[[Department]]", GetDepartment());
            }
            if (toReplace.Contains("[[Phone Number]]"))
            {
                toReplace = toReplace.Replace("[[Phone Number]]", GetPhone());
            }
            return toReplace;
        }
        private void EnableForm(bool enabled)
        {
            foreach (Control c in this.Controls)
            {
                c.Enabled = enabled;
            }
        }
        private void SendFinished(object sender, RunWorkerCompletedEventArgs e)
        {
            EnableForm(true);
            try
            {
                if (e.Result.ToString() == null || string.IsNullOrEmpty(e.Result.ToString().Trim()))
                {
                    WriteLog("Emails sent without errors!");
                }
                else
                {
                    WriteLog("Emails sent with errors:" + Environment.NewLine + Environment.NewLine + e.Result.ToString());
                    WriteErrors(e.Result.ToString());
                }
                _errors = "";
            } catch (Exception)
            {
                WriteLog("Emails could not be sent! This might be because you don't have access to the mailbox you are trying to send from!");
            }
        }

        private string GetName()
        {
            if (string.IsNullOrEmpty(_firstName))
            {
                while (!_autoDiscoverFinished)
                {
                    Thread.Sleep(100);
                }
                try
                {
                    _firstName = _service.ResolveName(_userEmailAddress.Substring(0, _userEmailAddress.IndexOf('@')), ResolveNameSearchLocation.DirectoryOnly, true)[0].Contact.DisplayName;
                }
                catch (Exception)
                {
                    while (_firstName == null || String.IsNullOrEmpty(_firstName.Trim()))
                    {
                        ModalPrompt firstNamePrompt = new ModalPrompt("Name Required", "Name not found! Please enter your name.");
                        firstNamePrompt.ShowDialog();
                        _firstName = firstNamePrompt.Result;
                    }
                }
            }
            return _firstName;
        }

        private string GetJob()
        {
            if (string.IsNullOrEmpty(_jobTitle))
            {
                while (!_autoDiscoverFinished)
                {
                    Thread.Sleep(100);
                }
                try
                {
                    _jobTitle = _service.ResolveName(_userEmailAddress.Substring(0, _userEmailAddress.IndexOf('@')), ResolveNameSearchLocation.DirectoryOnly, true)[0].Contact.JobTitle;
                }
                catch (Exception)
                {
                    while (_jobTitle == null || String.IsNullOrEmpty(_jobTitle.Trim()))
                    {
                        ModalPrompt jobPrompt = new ModalPrompt("Job Title Required", "Job title not found! Please enter your job title.");
                        jobPrompt.ShowDialog();
                        _jobTitle = jobPrompt.Result;
                    }
                }
            }
            return _jobTitle;
        }
        private string GetPhone()
        {
            if (string.IsNullOrEmpty(_phoneNumber))
            {
                while (!_autoDiscoverFinished)
                {
                    Thread.Sleep(100);
                }
                try
                {
                    _phoneNumber = _service.ResolveName(_userEmailAddress.Substring(0, _userEmailAddress.IndexOf('@')), ResolveNameSearchLocation.DirectoryOnly, true)[0].Contact.PhoneNumbers[PhoneNumberKey.BusinessPhone];
                }
                catch (Exception)
                {
                    while (_phoneNumber == null || String.IsNullOrEmpty(_phoneNumber.Trim()))
                    {
                        ModalPrompt phonePrompt = new ModalPrompt("Phone number Required", "Phone number not found! Please enter your phone number.");
                        phonePrompt.ShowDialog();
                        _phoneNumber = phonePrompt.Result;
                    }
                }
            }
            return _phoneNumber;
        }
        private string GetDepartment()
        {
            if (string.IsNullOrEmpty(_department))
            {
                while (!_autoDiscoverFinished)
                {
                    Thread.Sleep(100);
                }
                try
                {
                    _department = _service.ResolveName(_userEmailAddress.Substring(0, _userEmailAddress.IndexOf('@')), ResolveNameSearchLocation.DirectoryOnly, true)[0].Contact.Department;
                }
                catch (Exception)
                {
                    while (_department == null || String.IsNullOrEmpty(_department.Trim()))
                    {
                        ModalPrompt departmentPrompt = new ModalPrompt("Department Required", "Department not found! Please enter your _department.");
                        departmentPrompt.ShowDialog();
                        _department = departmentPrompt.Result;
                    }
                }
            }
            return _department;
        }
    }
}
