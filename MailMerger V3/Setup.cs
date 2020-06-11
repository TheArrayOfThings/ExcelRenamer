﻿using System;
using System.Drawing;
using System.Text;
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
                string rtfSignature = "", referenceAliases = "", defaultFont = "";

                SettingsImporter.ImportSettings(ref rtfSignature, ref inboxesBox, ref referenceAliases, ref defaultFont);
                refAliasesBox.Text = referenceAliases;
                defaultFontBox.Text = defaultFont;
                signatureBox.Rtf = rtfSignature;
            } else
            {
                defaultFontBox.Text = @"Font Name=Calibri, Font Size=12";
            }
            UsefulTools.SetMergeFields(new[] {"Name", "Email", "Job Title", "Department", "Phone Number"}, insertMergeFieldToolStripMenuItem, signatureBox);
        }

        private void addButton_Click(object sender, EventArgs e)
        {
            ModalPrompt inboxToAddPrompt = new ModalPrompt("Inbox details", "Enter email address for inbox.");
            inboxToAddPrompt.ShowDialog();
            string inboxToAdd = inboxToAddPrompt.Result;
            if (inboxToAdd != null && !(String.IsNullOrEmpty(inboxToAdd)))
            {
                if (UsefulTools.IsValidEmail(inboxToAdd.Trim()))
                {
                    inboxesBox.Items.Add(inboxToAdd.Trim());
                    inboxesBox.SelectedItem = inboxesBox.Items[inboxesBox.Items.Count - 1];
                } else
                {
                    MessageBox.Show(@"The email address you entered is invalid.", @"Invalid Email Address");
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
            if (inboxesBox.Items.Count < 1)
            {
                MessageBox.Show(@"You need to add at least one inbox.", @"Inbox missing");
            } else if (signatureBox.Text.Trim().Equals(""))
            {
                MessageBox.Show(@"Signature is blank.", @"Blank signature");
            } else
            {
                this.DialogResult = DialogResult.OK;
                signatureBox.SelectAll();
                signatureBox.SelectionFont = new Font(UsefulTools.GetFontValue(defaultFontBox.Text, "Name"), float.Parse(UsefulTools.GetFontValue(defaultFontBox.Text, "Size")));
                ExportSettings();
                this.Close();
            }            
        }

        //MenuItems events

        private void insertImageToolStripMenuItem_Click(object sender, EventArgs e)
        {
            OpenFileDialog getImage = new OpenFileDialog
            {
                Multiselect = false,
                Filter = @"Image Files | *.png; *.jpeg; *.bmp; *.gif"
            };
            Image image;
            if (getImage.ShowDialog() == DialogResult.OK && File.Exists(getImage.FileName))
            {
                try
                {
                    image = Image.FromFile(getImage.FileName);
                    Clipboard.SetImage(image);
                    signatureBox.Paste();
                }
                catch (OutOfMemoryException)
                {
                    MessageBox.Show(@"Please select a valid image.", @"Invalid Image");
                }
            }
            else
            {
                MessageBox.Show(@"Please select a valid image.", @"Invalid Image");
            }
        }

        private void hyperlinkToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ModalPrompt linkPrompt = new ModalPrompt("Hyperlink", "Enter link for hyperlink");
            linkPrompt.ShowDialog();
            string link = linkPrompt.Result;
            if (link != null & String.IsNullOrEmpty(link))
            {
                return;
            }
            UsefulTools.ReplaceHighlightedRtf("\\cf3\\ul " + signatureBox.SelectedText + " <" + link + ">\\cf7\\ulnone", signatureBox);
        }

        private void copyToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Clipboard.SetData(DataFormats.Rtf, signatureBox.SelectedRtf);
        }

        private void pasteToolStripMenuItem_Click(object sender, EventArgs e)
        {
            signatureBox.SelectedRtf = Clipboard.GetData(DataFormats.Rtf).ToString();
        }

        private void ExportSettings()
        {
            UsefulTools.TrimRtf(signatureBox);
            string toWrite = "****INBOXSTART****" + Environment.NewLine;
            foreach (string inbox in inboxesBox.Items)
            {
                toWrite += inbox + Environment.NewLine;
            }
            toWrite += ("****INBOXEND****" + Environment.NewLine);
            toWrite += ("****SIGNATURESTART****" + Environment.NewLine);
            toWrite += (signatureBox.Rtf + Environment.NewLine);
            toWrite += ("****SIGNATUREEND****" + Environment.NewLine);
            toWrite += ("****REFALIASESSTART****" + Environment.NewLine);
            toWrite += (refAliasesBox.Text.Trim() + Environment.NewLine);
            toWrite += ("****REFALIASESEND****" + Environment.NewLine);
            toWrite += ("****DEFAULTFONTSTART****" + Environment.NewLine);
            toWrite += (defaultFontBox.Text.Trim() + Environment.NewLine);
            toWrite += ("****DEFAULTFONTEND****" + Environment.NewLine);
            FileStream settingsFile = File.Create("Settings.txt");
            byte[] toWriteBytes = new UTF8Encoding(true).GetBytes(toWrite);
            settingsFile.Write(toWriteBytes, 0, toWriteBytes.Length);
            settingsFile.Close();
        }

        private void selectFontButton_Click(object sender, EventArgs e)
        {
            FontDialog changeFont = new FontDialog();
            if (changeFont.ShowDialog() == DialogResult.OK)
            {
                string fullFont = changeFont.Font.ToString();
                defaultFontBox.Text = @"Font Name=" + UsefulTools.GetFontValue(fullFont, "Name") + @", Font Size=" + Math.Round(float.Parse(UsefulTools.GetFontValue(fullFont, "Size")));
                signatureBox.SelectAll();
                signatureBox.SelectionFont = new Font(UsefulTools.GetFontValue(defaultFontBox.Text, "Name"), float.Parse(UsefulTools.GetFontValue(fullFont, "Size")));
                signatureBox.Select(0, 0);
            }
        }
    }
}
