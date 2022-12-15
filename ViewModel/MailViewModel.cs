using ExcelProject.Model;
using ExcelProject.ViewModel.Commands;
using Prism.Events;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Input;
using System.Xml.Linq;
using static ExcelProject.Model.Mail;
using Outlook = Microsoft.Office.Interop.Outlook;
namespace ExcelProject.ViewModel
{
    public class MailViewModel
    {
        private IEventAggregator _eventAggregator;
        private Mail _mail;
        public bool canExecuteSend => true;
        public ICommand OkCommand
        {
            get;
            set;

        }
        public Mail Mail
        {
            get
            {
                return _mail;
            }
            set
            {
                _mail = value;
            }
        }
        public MailViewModel(IEventAggregator eventAggregator)
        {
            this._eventAggregator = eventAggregator;
            Mail = new Mail();
            OkCommand = new RelayCommand(Send, () => canExecuteSend);
        }
        public void Send()
        {
            _eventAggregator.GetEvent<MailTransferEvent>().Publish(_mail);

        }
    }
}
