using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace IPA_GenerateReportExcel_Health.Models
{
    public class MailConfigModel
    {
        public string MailTo { get; set; }
        public string MailFrom { get; set; }
        public string MailCc { get; set; }
        public string Subject { get; set; }
        public string Body { get; set; }
        public string Attachment { get; set; }
        public string Smtp { get; set; }
    }
}
