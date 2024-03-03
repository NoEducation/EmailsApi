using System;
using System.Collections.Generic;

namespace OutlookApi
{
    public class POPClientEmail
    {
        public int MessageNumber { get; set; }
        public string From { get; set; }
        public string Subject { get; set; }
        public string Body { get; set; }
        public DateTime DateSent { get; set; }

        public ICollection<Attachment> Attachments { get; set; } = new HashSet<Attachment>();
    }
}
