using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;

namespace MailMerger_V3
{
    public partial class main : Form
    {
        Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
        OpenFileDialog openBook = new OpenFileDialog();
        Recipient[] recipients = new Recipient[100];
        int recipientNumber = 0;
        Microsoft.Office.Interop.Excel.Workbook importBook;
        Microsoft.Office.Interop.Excel.Worksheet importSheet;
        public main()
        {
            InitializeComponent();
            logBox.SelectAll();
            logBox.SelectionAlignment = HorizontalAlignment.Center;
        }

        private void importButton_Click(object sender, EventArgs e)
        {
            if (openBook.ShowDialog() == DialogResult.OK)
            {
                importBook = excel.Workbooks.Open(openBook.FileName);
                importSheet = importBook.Worksheets.get_Item(0);
                for (int i = 1; i < importSheet.Rows.Count; ++i)
                {
                    importSheet.Cells[i, 1].ToString();
                }
                writeLog("Import successful!");
            } else
            {
                writeLog("Import failed...");
            }
        }

        private void previewButton_Click(object sender, EventArgs e)
        {

        }

        private void sendButton_Click(object sender, EventArgs e)
        {

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

        private void previousButton_Click(object sender, EventArgs e)
        {

        }

        private void nextButton_Click(object sender, EventArgs e)
        {

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
        private void doubleRecipients()
        {
            Recipient[] newArray = new Recipient[recipients.Length * 2];
            for (int i = 0; i < recipients.Length; ++i)
            {
                newArray[i] = recipients[i];
            }
            recipients = newArray;
        }
    
    }
}
