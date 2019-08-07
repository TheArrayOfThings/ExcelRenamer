using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MailMerger_V3
{
    class Attachment
    {
        public string filePath { set; get; }
        public string fileName { set; get; }
        public Attachment(string fullPath)
        {
            filePath = fullPath;
            fileName = fullPath.Substring(fullPath.LastIndexOf("\\") + 1);
        }
    }
}
