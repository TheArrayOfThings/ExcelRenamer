using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;

namespace MailMerger_V3
{
    public partial class Setup : Form
    {
        public Setup()
        {
            InitializeComponent();
            if(File.Exists("Settings.txt"))
            {
                string rtfSignature = "";
                SettingsImporter.importSettings(ref rtfSignature, ref inboxesBox);
                signatureBox.Rtf = rtfSignature;
            }
            UsefulTools.setMergeFields(new string[] { "Name", "Email", "Job Title" }, insertMergeFieldToolStripMenuItem, signatureBox);
        }

        private void addButton_Click(object sender, EventArgs e)
        {
            string inboxToAdd = Prompt.ShowDialog("Enter email address for inbox.", "Inbox details");
            if (!(inboxToAdd.Trim().Equals("")))
            {
                if (UsefulTools.isValidEmail(inboxToAdd.Trim()))
                {
                    inboxesBox.Items.Add(inboxToAdd.Trim());
                    inboxesBox.SelectedItem = inboxesBox.Items[inboxesBox.Items.Count - 1];
                } else
                {
                    MessageBox.Show("The email address you entered is invalid.", "Invalid Email Address");
                }
            }
        }

        private void removeButton_Click(object sender, EventArgs e)
        {
            if (inboxesBox.SelectedIndex != -1)
            {
                inboxesBox.Items.RemoveAt(inboxesBox.SelectedIndex);
                if (inboxesBox.Items.Count > 0)
                {
                    inboxesBox.SelectedIndex = 0;
                }
            }
            
        }

        private void submitButton_Click(object sender, EventArgs e)
        {
            if (inboxesBox.Items.Count > 0 && (!(signatureBox.Text.Trim().Equals(""))))
            {
                this.DialogResult = DialogResult.OK;
                exportSettings();
                this.Close();
            } else
            {
                MessageBox.Show("You need to add at least one inbox.", "Inbox missing");
            }
            
        }
        private void copyToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Clipboard.SetData(DataFormats.Rtf, signatureBox.SelectedRtf);
        }

        private void hyperlinkToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string link = Prompt.ShowDialog("Enter link for hyperlink", "Hyperlink");
            UsefulTools.replaceHighlightedRtf("\\cf3\\ul " + signatureBox.SelectedText + " <" + link + ">\\cf7\\ulnone", signatureBox);
        }
        private void exportSettings()
        {
            UsefulTools.trimRtf(signatureBox);
            string toWrite = "";
            toWrite = "****INBOXSTART****" + Environment.NewLine;
            foreach (string inbox in inboxesBox.Items)
            {
                toWrite += inbox + Environment.NewLine;
            }
            toWrite += ("****INBOXEND****" + Environment.NewLine);
            toWrite += ("****SIGNATURESTART****" + Environment.NewLine);
            toWrite += (signatureBox.Rtf);
            toWrite += ("****SIGNATUREEND****" + Environment.NewLine);
            FileStream settingsFile = File.Create("Settings.txt");
            byte[] toWriteBytes = new UTF8Encoding(true).GetBytes(toWrite);
            settingsFile.Write(toWriteBytes, 0, toWriteBytes.Length);
            settingsFile.Close();
        }
    }
}
