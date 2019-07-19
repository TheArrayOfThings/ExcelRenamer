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

namespace MailMerger_V3
{
    public partial class main : Form
    {
        Excel.Application excel = new Excel.Application();
        OpenFileDialog openBook = new OpenFileDialog();
        int forenameColumn = -1, referenceColumn = -1, emailColumn = -1, currentRecipient = 0;
        Excel.Workbook importBook;
        Excel.Worksheet importSheet;
        Excel.Range usedRange;
        string[] columns;
        string[][] allData { set; get; }
        Boolean importSuccess = false;

        //EWS stuff
        ExchangeService service = new ExchangeService();
        bool autoDiscoverFinished = false;
        Thread startAutoDiscover;
        EmailAddress sendFrom;
        FolderId sentBoxSentItems;
        FolderId sentBoxDrafts;



        public main()
        {
            InitializeComponent();
            logBox.SelectAll();
            logBox.SelectionAlignment = HorizontalAlignment.Center;
            //Import settings etc
            inboxes.Items.Add("rflanagan@bournemouth.ac.uk");
            inboxes.Items.Add("anotheremail.com");
            inboxes.Items.Add("onlineservicesdev@bournemouth.ac.uk");
            inboxes.SelectedItem = inboxes.Items[0];
            //EWS setup
            service.UseDefaultCredentials = true;
            startAutoDiscover = new Thread(autoDiscoverThread);
            startAutoDiscover.IsBackground = true;
            startAutoDiscover.Start();
        }

        private void importButton_Click(object sender, EventArgs e)
        {
            if (openBook.ShowDialog() == DialogResult.OK)
            {
                importWorkbook(openBook.FileName);
                displayRecipient(0);
                importBook.Close();
            } else
            {
                writeLog("Please select an excel file.");
            }
        }

        private void previewButton_Click(object sender, EventArgs e)
        {

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
                sendAll();
            }
        }

        private void templateButton_Click(object sender, EventArgs e)
        {

        }

        private void setupButton_Click(object sender, EventArgs e)
        {

        }

        private void settingsButton_Click(object sender, EventArgs e)
        {

        }

        private void addButton_Click(object sender, EventArgs e)
        {

        }

        private void removeButton_Click(object sender, EventArgs e)
        {

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
                    if (findMatch(allData[0][i], new string[] { "forename", "firstname", "first name", "fore name" }))
                    {
                        forenameColumn = i;
                        writeLog("Forename is column " + (forenameColumn + 1));
                    }
                }
                if (referenceColumn == -1)
                {
                    if (findMatch(allData[0][i], new string[] { "reference", " id" }))
                    {
                        referenceColumn = i;
                        writeLog("Reference is column " + (referenceColumn + 1));
                    }
                }
                if (emailColumn == -1)
                {
                    if (findMatch(allData[0][i], new string[] { "email", " e-mail" }))
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
            }
        }
        private Boolean findMatch(string toCheck, string[] matchArray)
        {
            foreach (string eachString in matchArray)
            {
                if (toCheck.ToLower().Contains(eachString.ToLower()))
                {
                    return true;
                }
            }
            return false;
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
        private EmailMessage createEmail(string[] recipientData)
        {
            EmailMessage email = new EmailMessage(service);
            string htmlBody = RtfToHtmlConverter.RtfToHtml(bodyBox.Rtf).Html;
            Console.WriteLine(htmlBody);
            int imageOccurences = StringTools.CountStringOccurrences(htmlBody, "<img");
            int lastImgOccurence = 0;
            for (int i = 0; i < imageOccurences; ++i)
            {
                int imageIndex = htmlBody.IndexOf("<img");
                int srcIndex = htmlBody.IndexOf("scr=\"", imageIndex);
                int srcEnd = htmlBody.IndexOf("/>", srcIndex);
                int fileNameLength = srcEnd - srcIndex;
                string fileLocation = htmlBody.Substring(srcIndex, srcEnd);

                //email.Attachments.AddFileAttachment(htmlBody.Substring(srcIndex, srcEnd));
                //Find img src, attach image to email, replace src with <img src="cid:yourcontentid" />
            }
            email.Body = htmlBody;
            email.Subject = subjectBox.Text;
            email.ToRecipients.Add(recipientData[emailColumn]);
            email.From = sendFrom;
            return email;

        }
        private void sendAll()
        {
            string sendingInbox = inboxes.SelectedItem.ToString().Trim();
            Mailbox sentBox = new Mailbox(sendingInbox);
            sentBoxSentItems = new FolderId(WellKnownFolderName.SentItems, sentBox);
            sentBoxDrafts = new FolderId(WellKnownFolderName.Drafts, sentBox);
            sendFrom = new EmailAddress(sendingInbox);
            for (int i = 1; i < allData.Length; ++i)
            {
                sendEmail(allData[i]);
            }
        }
        private void sendEmail(String[] recipientData)
        {
            EmailMessage finalEmail = createEmail(recipientData);
            finalEmail.Save(sentBoxDrafts);
            finalEmail.SendAndSaveCopy(sentBoxSentItems);
        }
    }
}
