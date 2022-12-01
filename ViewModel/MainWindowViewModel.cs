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

namespace ExcelProject.ViewModel;

public class MainWindowViewModel : INotifyPropertyChanged
{
    public ICommand ShowCommand
    {
        get;
        set;

    }
    public ICommand AddCommand
    {
        get;
        set;

    }

    public ICommand ImportCommand
    {
        get;
        set;
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
    public MainWindowViewModel()
    {
        ShowCommand = new RelayCommand(OnShow, () => canExecuteOnShow);
        ImportCommand = new RelayCommand(Import, () => canExecuteImport);
        AddCommand = new RelayCommand(Add, () => canExecuteAdd);
    }
    public bool canExecuteOnShow => true;
    public bool canExecuteImport => true;
    public bool canExecuteAdd => true;
    private string _fileName;

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
    public bool canExecute => true;

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
  
    public void Import()
    {
        HSSFWorkbook wb;
        HSSFSheet sh;
        String Sheet_name;

        using (var fs = new FileStream(fileName, FileMode.Open, FileAccess.Read))
        {
            wb = new HSSFWorkbook(fs);

            Sheet_name = wb.GetSheetAt(0).SheetName;  
        }
      
        
        

        
        sh = (HSSFSheet)wb.GetSheet(Sheet_name);
        dataTable = new DataTable(sh.SheetName);

    }

    

    public void Add()
    {
        AddEmployee add = new AddEmployee();
        add.Show();
    }
}
