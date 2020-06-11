namespace MailMerger_V3
{
    class Attachment
    {
        public string FilePath { set; get; }
        public string FileName { set; get; }
        public Attachment(string fullPath)
        {
            FilePath = fullPath;
            FileName = fullPath.Substring(fullPath.LastIndexOf("\\") + 1);
        }
    }
}
