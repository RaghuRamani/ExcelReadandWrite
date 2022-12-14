using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.OleDb;
using System.Collections.ObjectModel;
using Prism.Events;
using System.Security.Policy;
using System.ComponentModel;
using static ExcelProject.Model.EmployeeTransferEvent;

namespace ExcelProject.Model
{
    public class EmployeeTransferEvent : PubSubEvent<Employee>
    {
    }
    public class Employee : IDataErrorInfo
    {
        public string Error
        {
            get
            {
                return string.Empty;
            }
        }

        public string this[string columnName]
        {
            get
            {
                string result = String.Empty;
                if (columnName == "EmpName")
                {
                    if (EmpName.Length < 2 || EmpName.Length > 12)
                    {
                        result = "Name should be between range 2-12";
                    }
                }

                return result;
            }
        }

        public int EmpNo 
        {
            get;
            set;
        }
        public string EmpName
        {
            get;
            set;

        } = "";
        public int Salary 
        {
            get;
            set;
        }
        public string DeptName 
        {
            get;
            set;
        }

     
    }
}
