using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace IPA_GenerateReportExcel_Health.Models
{
    public class ConfigModel
    {
       public string ExcelFilePath { get; set; }
       public string MailBody { get; set; }
       public string MailCc { get; set; }
       public string MailFrom { get; set; }
       public string MailSmtp { get; set; }
       public string MailStatus { get; set; }
       public string MailSubject { get; set; }
       public string MailTo { get; set; }
       public string UplaodDirectory { get; set; }
       public string UplaodFilePath { get; set; }
       public string UplaodHost { get; set; }
       public string UplaodPassword { get; set; }
       public string UplaodPort { get; set; }
       public string UplaodStatus { get; set; }
       public string UplaodUsername { get; set; }
    }
}
