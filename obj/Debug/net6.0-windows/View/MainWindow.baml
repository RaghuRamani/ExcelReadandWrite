   M S B A M L   `   `   `  ????G  CExcelProject, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null?$clr-namespace:ExcelProject.ViewModelExcelProject.ViewModel  5   	   P LPrism, Version=8.1.97.5141, Culture=neutral, PublicKeyToken=40ee6c3a2184dc59:)clr-namespace:Prism.Events;assembly=PrismPrism.Events 5   	       ExcelProject.MainWindow   5   	   W SPresentationCore, Version=6.0.2.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35\ XPresentationFramework, Version=6.0.2.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35e aSystem.Windows.Controls.Ribbon, Version=6.0.2.0, Culture=neutral, PublicKeyToken=b77a5c561934e089R NWindowsBase, Version=6.0.2.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35F 9http://schemas.microsoft.com/winfx/2006/xaml/presentation     5   	   R NSystem.Xaml, Version=6.0.2.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     5   	   @d2http://schemas.microsoft.com/expression/blend/2008     5   	   Jmc;http://schemas.openxmlformats.org/markup-compatibility/2006     5   	   6local$clr-namespace:ExcelProject.ViewModel     5   	   <system)clr-namespace:Prism.Events;assembly=Prism     5   	     ? Title$  
MainWindow??5	   	   $	??800??5
   	   
???????$
??1000??5   	   .??5      ?  ? RowDefinitions 5)      ?? 5*      $
??272* ?6   ?? 5+      $	??71* ?6   .+?5-      ?? $
??Left=?6   $??	23,21,0,0q?5.      
??????? ?? TextWrapping$
 Wrap=?50      $	??Top=?51      $	??834??52      $??28??54      ??53      ?? *
fileName+54      ?? 55      $??Import??6   $
??Left=?56      $??20(?57      
???????$??Times New Roman)?59      $??	ExtraBold"?5:      $??869,112,0,0q?5;      $	??Top=?5<      $??39??5=      $	??118??5>      ??5?      ?? *ImportCommand+?? 5@      $??
Add Record??6   $
??Left=?5A      $??868,158,0,0q?5B      $??20(?5C      
???????$??Times New Roman)?5F      $??	ExtraBold"?5G      $	??Top=?5H      $??40??5I      $	??120??5J      ??5E      ?? *
AddCommand+5J      ?? 5K      $??Browse??6   $
??Left=?5L      
???????$??
867,66,0,0q?5N      $??20(?5O      $??Times New Roman)?5P      $??	ExtraBold"?5Q      $	??Top=?5S      $??39??5T      $	??120??5U      ??5R      ?? *ShowCommand+5U      &   System.Windows.Controls.DataGrid  5V      -   5        $??DataGrid??5X      
???????$??30,62,145,41q?5W         AutoGenerateColumns$ False??5Z      $??2??5[      ??5Y      ?? *dataTable.DefaultView+   Columns 5\      0  *System.Windows.Controls.DataGridTextColumn  5^   
   ,  &System.Windows.Controls.DataGridColumn   Header$ Employee No.??6   1  +System.Windows.Controls.DataGridBoundColumn   Binding 5_      ?? *EmpNo+  5`   
   $ Employee Name??6    5a      ?? *	EmpName+  5b   
   $ Salary??6    5c      ?? *Salary+  5d   
   $ Department Name??6    5e      ?? *
DeptName+5h      ?? 5i      $??Export??6   $
??Left=?5j      $??
868,19,0,0q?5k      
???????$??20(?5m      $??Times New Roman)?5o      $??	ExtraBold"?5p      $	??Top=?5q      $??39??5r      $	??119??5s      ??5n      ?? *ExportCommand+5s      ?? 5t      $??	Send Mail??6   $
??Left=?5u      $??868,205,0,0q?5v      $??20(?5w      
???????$??Times New Roman)?5z      $??	ExtraBold"?5{      $	??Top=?5|      $??40??5}      $	??120??5~      ??5y      ?? *MailCommand+5~      5?      5?      