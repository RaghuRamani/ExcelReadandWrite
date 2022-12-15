using Microsoft.Office.Interop.Outlook;
using Prism.Events;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelProject.Model
{
    public  class Mail
    {

        public class MailTransferEvent : PubSubEvent<Mail>
        {
        }
        public String MailTo
        {
            get;
            set;
        }
        public Attachments Attachments
        {
            get;
            set;

        } 
        public string MailMessage
        {
            get;
            set;
        }
        public string Subject
        {
            get;
            set;
        }
        public string cc
        {
            get;
            set;
        }
    }
}
