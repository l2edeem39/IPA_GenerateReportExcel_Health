using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace IPA_GenerateReportExcel_Health.Models
{
    public class SFTPConfigModel
    {
        public string FileToUpload { get; set; }
        public string Host { get; set; }
        public string Port { get; set; }
        public string Username { get; set; }
        public string Password { get; set; }
        public string FileDirectory { get; set; }
        public string Status { get; set; }
    }
}
