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
        }

        private void addButton_Click(object sender, EventArgs e)
        {
            string inboxToAdd = Prompt.ShowDialog("Enter email address for inbox.", "Inbox details");
            if (!(inboxToAdd.Trim().Equals("")))
            {
                inboxesBox.Items.Add(inboxToAdd.Trim());
                inboxesBox.SelectedItem = inboxesBox.Items[inboxesBox.Items.Count - 1];
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
        private void exportSettings()
        {
            StringTools.trimRtf(signatureBox);
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
