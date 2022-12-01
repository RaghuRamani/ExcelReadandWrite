using ExcelProject.Model;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Data.OleDb;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelProject.ViewModel
{
    public class DataAccess
    {
        OleDbConnection Conn;
        OleDbCommand Cmd;

        public DataAccess()
        {
            Conn = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\\Users\\Raghu Ramani\\Downloads\\USELESS\\Employee.xlsx;Extended Properties=\"Excel 12.0 Xml;HDR=YES;\"");
             async Task<ObservableCollection<Employee>> GetDataFormExcelAsync()
            {
                ObservableCollection<Employee> Employees = new ObservableCollection<Employee>();
                await Conn.OpenAsync();
                Cmd = new OleDbCommand();
                Cmd.Connection = Conn;
                Cmd.CommandText = "Select * from [Sheet1$]";
                var Reader = await Cmd.ExecuteReaderAsync();
                while (Reader.Read())
                {
                    Employees.Add(new Employee()
                    {
                        EmpNo = Convert.ToInt32(Reader["EmpNo"]),
                        EmpName = Reader["EmpName"].ToString(),
                        DeptName = Reader["DeptName"].ToString(),
                        Salary = Convert.ToInt32(Reader["Salary"])
                    });
                }
                Reader.Close();
                Conn.Close();
                return Employees;
            }
        }
    }
}
