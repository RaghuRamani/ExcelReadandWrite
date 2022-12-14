using ExcelProject.Model;
using ExcelProject.ViewModel.Commands;
using NPOI.SS.Formula.Functions;
using Prism.Events;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Input;

namespace ExcelProject.ViewModel
{
    
    public class AddEmployeeViewModel
    {
        public ICommand SubmitCommand
        {
            get;
            set;

        }
        public bool canExecuteSave => true;
        private IEventAggregator _eventAggregator;
        private Employee _employee;
        public AddEmployeeViewModel(IEventAggregator eventAggregator)
        {
            this._eventAggregator = eventAggregator;
            Employee = new Employee();
            SubmitCommand = new RelayCommand(save, () => canExecuteSave);
        }

        public void save()
        {
            _eventAggregator.GetEvent<EmployeeTransferEvent>().Publish(_employee);


        }

        public Employee Employee
        {
            get
            {
                return _employee;
            }
            set
            {
                _employee = value;
            }
        }

    }
}
