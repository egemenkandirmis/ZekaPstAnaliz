using System;
using System.Collections.Generic;
using System.Text;

namespace MailBackupForm.Properties
{
    public class Mail
    {
        public DateTime DeliveryTime { get; set; }
        public string SenderEmailAddress { get; set; }
        public string Subject { get; set; }
        public string DisplayTo { get; set; }
        public string DisplayBcc { get; set; }
        public string DisplayCc { get; set; }
        public string Body { get; set; }
        public string FilePath { get; set; }

    }
}
