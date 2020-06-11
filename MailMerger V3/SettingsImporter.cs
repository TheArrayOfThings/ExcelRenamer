using System;
using System.Windows.Forms;
using System.IO;

namespace MailMerger_V3
{
    static class SettingsImporter
    {
        public static void ImportSettings(ref string rtfSignature, ref ComboBox inboxes, ref string refAliases, ref string defaultFont)
        {
            string[] settingsContents = File.ReadAllLines("Settings.txt");
            for (int i = 0; i < settingsContents.Length; ++i)
            {
                if (settingsContents[i].Equals("****INBOXSTART****"))
                {
                    ++i;
                    //Console.WriteLine("Starting inbox import...");
                    while (i <= settingsContents.Length && (!(settingsContents[i].Equals("****INBOXEND****"))))
                    {
                        inboxes.Items.Add(settingsContents[i]);
                        ++i;
                    }

                }
                if (settingsContents[i].Equals("****SIGNATURESTART****"))
                {
                    ++i;
                    //Console.WriteLine("Starting signature import...");
                    rtfSignature = "";
                    while (i <= settingsContents.Length && (!(settingsContents[i].Equals("****SIGNATUREEND****"))))
                    {
                        rtfSignature += settingsContents[i] + Environment.NewLine;
                        ++i;
                    }
                }
                if (settingsContents[i].Equals("****REFALIASESSTART****"))
                {
                    ++i;
                    //Console.WriteLine("Starting reference aliases import...");
                    refAliases = settingsContents[i].Trim();
                }
                if (settingsContents[i].Equals("****DEFAULTFONTSTART****"))
                {
                    ++i;
                    //Console.WriteLine("Starting reference aliases import...");
                    defaultFont = settingsContents[i].Trim();
                }
            }
            if (inboxes.Items.Count > 0)
            {
                inboxes.SelectedItem = inboxes.Items[0];
            }
        }
    }
}
