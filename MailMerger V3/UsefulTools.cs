using System;
using System.Windows.Forms;
using System.Net.Mail;

namespace MailMerger_V3
{
    public static class UsefulTools
    {
        // Count occurrences of strings.
        public static int CountStringOccurrences(string text, string pattern)
        {
            // Loop through all instances of the string 'text'.
            int count = 0;
            int i = 0;
            while ((i = text.IndexOf(pattern, i)) != -1)
            {
                i += pattern.Length;
                count++;
            }
            return count;
        }
        public static void ReplaceHighlightedRtf(string toAdd, RichTextBox toReplace)
        {
            toReplace.SelectedRtf = "{\\rtf1" + toAdd + "}}";
        }
        public static void SetMergeFields(String[] toSet, ToolStripMenuItem mainMenu, RichTextBox relevantBox)
        {
            foreach (string eachString in toSet)
            {
                ToolStripMenuItem newItem = new ToolStripMenuItem(eachString);
                newItem.Click += new System.EventHandler((sender, e) => InsertMergeField(sender, relevantBox));
                mainMenu.DropDownItems.Add(newItem);
            }
        }
        private static void InsertMergeField(object sender, RichTextBox relevantBox)
        {
            relevantBox.SelectedRtf = "{\\rtf1" + "[[" + sender + "]]" + "}}";
        }
        public static string GrabImgSrc(int startIndex, string toGrab)
        {
            int srcIndex = toGrab.IndexOf("src=", startIndex) + 5;
            int srcEnd = toGrab.IndexOf(" />", srcIndex);
            int imgTagLength = (srcEnd - srcIndex) - 1;
            return toGrab.Substring(srcIndex, imgTagLength);
        }
        public static bool FindMatch(string toCheck, string[] matchArray)
        {
            foreach (string eachString in matchArray)
            {
                if (eachString.Trim().Equals(""))
                {
                    return false;
                }
                if (toCheck.ToLower().Contains(eachString.ToLower()))
                {
                    return true;
                }
            }
            return false;
        }
        public static string ReplaceFirst(string searchText, string search, string replace)
        {
            int pos = searchText.IndexOf(search);
            if (pos < 0)
            {
                return searchText;
            }
            return searchText.Substring(0, pos) + replace + searchText.Substring(pos + search.Length);
        }
        public static void TrimRtf(RichTextBox toTrim)
        {
            while (toTrim.Text.EndsWith("\n"))
            {
                toTrim.Select(toTrim.Text.LastIndexOf("\n"), 2);
                toTrim.SelectedRtf = "";
            }
            while (toTrim.Text.StartsWith("\n"))
            {
                toTrim.Select(toTrim.Text.IndexOf("\n"), 1);
                toTrim.SelectedRtf = "";
            }
        }
        public static bool IsValidEmail(string toCheck)
        {
            try
            {
                MailAddress test = new MailAddress(toCheck);
                return true;
            }
            catch (FormatException)
            {
                return false;
            }
        }
        public static string GetFontValue(string toParse, string attributeName)
        {
            int startIndex = toParse.IndexOf(attributeName) + attributeName.Length + 1;
            int endIndex;
            if (toParse.IndexOf(';', startIndex) == -1)
            {
                endIndex = toParse.Length;
            }
            else
            {
                endIndex = toParse.IndexOf(';', startIndex);
            }
            int length = endIndex - startIndex;
            return toParse.Substring(startIndex, length);
        }
    }
}