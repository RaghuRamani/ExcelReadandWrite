using ExcelProject.ViewModel.Commands;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data.OleDb;
using System.Data;
using System.IO.Enumeration;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Input;
using System.Windows.Interop;
using NPOI.POIFS.NIO;
using ExcelProject.Model;
using ExcelProject.View;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.IO;
using System.Collections;
using NPOI.SS.Formula.Functions;
using NPOI.Util;
using NPOI.SS.Formula.Atp;
using Prism.Events;
using System.Windows.Controls;
using NPOI.SS.Formula.Eval;
using NPOI.SS.Formula;
using Prism.Services.Dialogs;
using excelfile= Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Excel;
using DataTable = System.Data.DataTable;
using static ExcelProject.Model.Mail;
using System.Xml.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
namespace ExcelProject.ViewModel;

public class MainWindowViewModel : INotifyPropertyChanged
{
    protected readonly IEventAggregator _eventAggregator;
    private DataRow _row;
    public ICommand ShowCommand
    {
        get;
        set;

    }
    public ICommand ExportCommand
    {
        get;
        set;

    }
    public ICommand MailCommand
    {
        get;
        set;

    }
    /*public DataRow row
    {
        get;
        set;
    }*/

    public ICommand AddCommand
    {
        get;
        set;

    }
    public ICommand EnterCommand
    {
        get;
        set;

    }

    public ICommand ImportCommand
    {
        get;
        set;
    }
    public Employee Employee
    {
        get;
        set;

    }
    public Mail Mail
    {
        get;
        set;

    }
    public DataRow row
    {
        get
        {
            return _row;
        }
        set
        {
            _row = value;
            OnPropertyChanged("row");
        }
    }
    public DataTable _dataTable;
    public DataTable dataTable
    {

        get
        {
            return _dataTable;
        }
        set
        {
            _dataTable = value;
            OnPropertyChanged("dataTable");
        }
    }

    public string _addEmployee;
    public String AddEmployee
    {

        get
        {
            return _addEmployee;
        }
        set
        {
            _addEmployee = value;

        }
    }
    private string _fileName;
    private string _filePath;
    public string fileName
    {

        get
        {
            return _fileName;
        }
        set
        {
            _fileName = value;
            OnPropertyChanged("fileName");
        }
    }
    public string filePath
    {

        get
        {
            return _filePath;
        }
        set
        {
            _filePath = value;
            OnPropertyChanged("filePath");
        }
    }
    public MainWindowViewModel(IEventAggregator eventAggregator)
    {
        this._eventAggregator = eventAggregator;
        this._eventAggregator.GetEvent<EmployeeTransferEvent>().Subscribe((_employee) => { enter(this.Employee = _employee);  });
        this._eventAggregator.GetEvent<MailTransferEvent>().Subscribe((_mail) => { Send(this.Mail = _mail); });
        ShowCommand = new RelayCommand(OnShow, () => canExecuteOnShow);
        ImportCommand = new RelayCommand(Import, () => canExecuteImport);
        AddCommand = new RelayCommand(Add, () => canExecuteAdd);
        ExportCommand = new RelayCommand(export, () => canExecuteExport);
        MailCommand = new RelayCommand(MailTo, () => canExecuteMail);
    }

    public bool canExecute => true;
    public bool canExecuteOnShow => true;
    public bool canExecuteImport => true;
    public bool canExecuteAdd => true;
    public bool canExecuteEnter => true;
    public bool canExecuteExport => true;
    public bool canExecuteMail => true;





    private void OnShow()
    {
        var dialog = new Microsoft.Win32.OpenFileDialog();
        dialog.Filter = "Excel documents (.xls)|*.xls";



        bool? result = dialog.ShowDialog();


        if (result == true)
        {

            fileName = dialog.FileName;
        }
    }


    public event PropertyChangedEventHandler PropertyChanged;

    public void OnPropertyChanged(string name)
    {
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
    }
    /* public void Import()
     {
         OleDbConnection _Connection = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0; data source=" + fileName + ";Extended Properties=\"Excel 8.0;HDR=NO;IMEX=1;\"");
         _Connection.Open();

         OleDbDataAdapter theDataAdapter = new OleDbDataAdapter("SELECT * FROM [Sheet1$]", _Connection);

         DataSet ds = new DataSet();
         theDataAdapter.Fill(ds);
         DataTable dt = ds.Tables[0];

     }*/
    String Sheet_name;
    public void Import()
    {
        try
        {
            HSSFWorkbook wb;
            HSSFSheet sh;
            

            using (var fs = new FileStream(fileName, FileMode.Open, FileAccess.Read))
            {
                wb = new HSSFWorkbook(fs);

                Sheet_name = wb.GetSheetAt(0).SheetName;
            }
            sh = (HSSFSheet)wb.GetSheet(Sheet_name);
            dataTable = new DataTable(sh.SheetName);
            var headerRow = sh.GetRow(0);
            foreach (var headerCell in headerRow)
            {
                dataTable.Columns.Add(headerCell.ToString());
            }
            for (int i = 1; i < sh.PhysicalNumberOfRows; i++)
            {
                var sheetRow = sh.GetRow(i);
                if (sh.GetRow(0) == null)
                {
                    throw new Exception("Blank File Selected!!");
                }
                var dtRow = dataTable.NewRow();
                dtRow.ItemArray = dataTable.Columns
                    .Cast<DataColumn>()
                    .Select(c => sheetRow.GetCell(c.Ordinal, MissingCellPolicy.CREATE_NULL_AS_BLANK).ToString())
                    .ToArray();
                dataTable.Rows.Add(dtRow);

            }

        }
       
        catch (Exception exception)
        {

            MessageBox.Show(exception.ToString());
        }
    }
        
        
    



    public void Add()
    {
        AddEmployee add = new AddEmployee();
        add.Show();


    }

    /* public void save()
    {
         row = dataTable.NewRow();
         row["EmpNo"] = Employee.EmpNo;
         row["EmpName"] = Employee.EmpName;
         row["Salary"] = Employee.Salary;
         row["DeptName"] = Employee.DeptName;
         dataTable.Rows.Add(row);
    }
  */
    public void enter(object e)
    {   Employee employee= (Employee)e;
        row = dataTable.NewRow();
        row["EmpNo"] = Employee.EmpNo;
        row["EmpName"] = Employee.EmpName;
        row["Salary"] = Employee.Salary;
        row["DeptName"] = Employee.DeptName;
        dataTable.Rows.Add(row);
    }
    /* public void export()
     {
         SaveFileDialog save = new SaveFileDialog();
         save.Filter = "Excel File (*.xls)|*.xls|Show All Files (*.*)|*.*";
         bool? result = save.ShowDialog();


         if (result == true)
         {

             filePath = save.FileName;
         }
         DataTable dt = new DataTable();
         DataTable d = new DataTable();

         dt = dataTable.Copy();
         var f= new FileStream(filePath, FileMode.Append, FileAccess.Write);
         using (f)
         {

             IWorkbook workbook = new HSSFWorkbook();
             ISheet excelSheet = workbook.CreateSheet(dt.TableName);
             List<string> columns = new List<string>();
             IRow row = excelSheet.CreateRow(0);
             int columnIndex = 0;

             foreach (System.Data.DataColumn column in dt.Columns)
             {
                 columns.Add(column.ColumnName);
                 row.CreateCell(columnIndex).SetCellValue(column.ColumnName);
                 columnIndex++;
             }

             int rowIndex = 1;
             foreach (DataRow dsrow in dt.Rows)
             {
                 row = excelSheet.CreateRow(rowIndex);
                 int cellIndex = 0;
                 foreach (String col in columns)
                 {
                     row.CreateCell(cellIndex).SetCellValue(dsrow[col].ToString());
                     cellIndex++;

                 }

                 rowIndex++;

             }

             workbook.Write(f,true);
         }
     }
    */
    public void export()
    {
        excelfile.Application excel;
        excelfile.Workbook excelworkBook;
        excelfile.Worksheet excelSheet;
        excelfile.Range excelCellrange;

        try
        {
            excel = new excelfile.Application();
            excel.Visible = false;
            excel.DisplayAlerts = false;

            excelworkBook = excel.Workbooks.Add(Type.Missing);

            excelSheet = (excelfile.Worksheet)excelworkBook.ActiveSheet;
            excelSheet.Name = Sheet_name;

            int rowcount = 1;
            

            foreach (DataRow datarow in dataTable.Rows)
            {
                rowcount += 1;

                for (int i = 1; i <= dataTable.Columns.Count; i++)
                {
                    if (rowcount == 3)
                    {
                        excelSheet.Cells[1, i] = dataTable.Columns[i - 1].ColumnName;
                    }
                   
                    excelSheet.Cells[rowcount, i] = datarow[i-1].ToString();
                }


            }


            SaveFileDialog save = new SaveFileDialog();
            save.Filter = "Excel File (*.xls)|*.xls|Show All Files (*.*)|*.*";
            bool? result = save.ShowDialog();


            if (result == true)
            {

                filePath = save.FileName;
            }
            excelworkBook.SaveAs(filePath); ;
            excelworkBook.Close();
            excel.Quit();
            
        }
        catch (Exception ex)
        {
            MessageBox.Show(ex.Message);
            
        }
    }
    public void MailTo()
    {
        Email email = new Email();
        email.Show();
    }
   /* public void Send(object m)
    {
        Mail mail = (Mail)m;
        try
        {
            Outlook.Application outlookApp = new Outlook.Application();
            Outlook.MailItem mailItem = (Outlook.MailItem)outlookApp.CreateItem(Outlook.OlItemType.olMailItem);

            mailItem.Subject = Mail.Subject;
            mailItem.To = Mail.MailTo.Replace(',', ';');


            if (!string.IsNullOrEmpty(Mail.cc))
                mailItem.CC = Mail.cc;

            mailItem.Body = Mail.MailMessage;
            mailItem.Importance = Outlook.OlImportance.olImportanceHigh;

            mailItem.Send();
        }
        catch (Exception ex)
        {
            Console.WriteLine(ex);
        }
    }*/
    public void Send(object m)
    {
        Mail mail = (Mail)m;
        try
        {
            Outlook.Application outlookApp = new Outlook.Application();
            Outlook.MailItem mailItem = (Outlook.MailItem)outlookApp.CreateItem(Outlook.OlItemType.olMailItem);

            mailItem.Subject = Mail.Subject;
            mailItem.To = Mail.MailTo.Replace(',', ';');
            OpenFileDialog attachment = new OpenFileDialog();

            attachment.Title = "Select a file to send";
            attachment.ShowDialog();

            if (attachment.FileName.Length > 0)
            {
                Mail.Attachments.Add(
                    attachment.FileName,
                    Outlook.OlAttachmentType.olByValue,
                    1,
                    attachment.FileName);
                
            }

            if (!string.IsNullOrEmpty(Mail.cc))
                mailItem.CC = Mail.cc;

            mailItem.Body = Mail.MailMessage;
            mailItem.Importance = Outlook.OlImportance.olImportanceHigh;

           
        }
        catch (Exception ex)
        {
            Console.WriteLine(ex);
        }
    }
}
