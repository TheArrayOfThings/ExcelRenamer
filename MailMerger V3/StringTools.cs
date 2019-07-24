using System;
using System.Collections.Generic;
using System.Windows.Forms;

public static class StringTools
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
    public static string grabImgSRC(int startIndex, string toGrab)
    {
        int srcIndex = toGrab.IndexOf("src=", startIndex) + 4;
        int srcEnd = toGrab.IndexOf("/>", srcIndex);
        int imgTagLength = (srcEnd - srcIndex) - 2;
        return toGrab.Substring(srcIndex, imgTagLength);
    }
    public static string grabStyle(string toGrab)
    {
        try
        {
            int fontStartIndex = toGrab.IndexOf("font-family:");
            int fontEndIndex = toGrab.IndexOf("pt;", fontStartIndex);
            int fontLength = (fontEndIndex - fontStartIndex) + 3;
            return toGrab.Substring(fontStartIndex, fontLength);
        } catch (System.ArgumentOutOfRangeException)
        {
            return "";
        }
    }
    public static string replaceStyle(string toGrab, string toReplace)
    {
        //Finds first font family and font size from 'toGrab' and replaces the font styles in 'toReplace'
        //Limitation = cannot have differening font families or sizes in signature
        return toReplace.Replace(grabStyle(toReplace), grabStyle(toGrab));
    }
    public static bool findMatch(string toCheck, string[] matchArray)
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
    public static string ReplaceFirst(string searchText, string search, string replace)
    {
        int pos = searchText.IndexOf(search);
        if (pos < 0)
        {
            return searchText;
        }
        return searchText.Substring(0, pos) + replace + searchText.Substring(pos + search.Length);
    }
    public static void trimRtf(RichTextBox toTrim)
    {
        while(toTrim.Text.EndsWith("\n"))
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
}