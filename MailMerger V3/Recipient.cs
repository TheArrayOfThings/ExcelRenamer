using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MailMerger_V3
{
    class Recipient
    {
        string foreName { get; set; }
        string reference { get; set; }
        string email { get; set; }
        Recipient(string foreName, string reference, string email)
        {
            this.foreName = foreName;
            this.reference = reference;
            this.email = email;
        }
    }
}
