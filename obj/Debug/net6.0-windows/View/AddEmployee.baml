   M S B A M L   `   `   `  ????G  CExcelProject, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null?$clr-namespace:ExcelProject.ViewModelExcelProject.ViewModel  5   	   #    ExcelProject.View.AddEmployee   5   	   W SPresentationCore, Version=6.0.2.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35\ XPresentationFramework, Version=6.0.2.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35e aSystem.Windows.Controls.Ribbon, Version=6.0.2.0, Culture=neutral, PublicKeyToken=b77a5c561934e089R NWindowsBase, Version=6.0.2.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35F 9http://schemas.microsoft.com/winfx/2006/xaml/presentation     5   	   R NSystem.Xaml, Version=6.0.2.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     5   	   @d2http://schemas.microsoft.com/expression/blend/2008     5   	   Jmc;http://schemas.openxmlformats.org/markup-compatibility/2006     5   	   6local$clr-namespace:ExcelProject.ViewModel     5   	     ? Title$  AddEmployee??5   	   $	??450??5	   	   $	??800??5
   	    ? 	Resources 5       ?? 
TargetType ?? Triggers P? Property P? Value ?? Property ?? Value ?? Self	 ?? RelativeSource
 ?? Path??%?   '??       ?? 5      " ??6    5      P? 5   
    ?@=?HasError$
 True??5      . ?5      ??  ?@?ToolTip 5      ?? #	 Z $&
  (Validation.Errors)/ErrorContent?5      5      5      .??5      ? .+?5      ?? $
??Left=?6   $??345,163,0,0q?5       ?? TextWrapping$
 Wrap=?5      $	??Top=?5       $	??182??5!       J? RenderTransformOrigin$ 0.493,0.353&?5"      $??29??5#      ??5      ?? *Employee.EmpNo+5#      ?? 5$      $
??Left=?6   $??345,225,0,0q?5%      $
 Wrap=?5&      $	??Top=?5(      $	??182??5)      $ 0.493,0.353&?5*      $??29??5+      ??5'      ?? *Employee.EmpName+ ?? ValidatesOnDataErrors$
 True??5+      ?? 5,      $
??Left=?6   $??344,290,0,0q?5-      $
 Wrap=?5.      $	??Top=?50      $	??182??51      $ 0.493,0.353&?52      $??29??53      ??5/      ?? *Employee.Salary+53      ?? 54      $
??Left=?6   $??218,166,0,0q?55      $??Employee No.:-??56      $
??Wrap=?57      $	??Top=?58      $??19??59      $??94??5:      ?? 5;      $
??Left=?6   $??214,226,0,0q?5<      $
??Wrap=?5=      $??Employee Name :-??5>      $	??Top=?5?      $??28??5@      $	??123??5A      ?? 5B      $
??Left=?6   $??216,298,0,0q?5C      $
??Wrap=?5D      $??	Salary :-??5E      $	??Top=?5F      $??24??5G      $??74??5H      ?? 5I      $
??Left=?6   $??343,338,0,0q?5J      $
 Wrap=?5K      $	??Top=?5L      $	??182??5N      $ 0.493,0.353&?5O      $??29??5P      ??5M      ?? *Employee.DeptName+5P      ?? 5Q      $
??Left=?6   $??215,346,0,0q?5R      $
??Wrap=?5S      $??Department Name:-??5T      $	??Top=?5U      $??24??5V      $	??115??5W      ?? 5X      $??Submit??6   $
??Left=?5Y      $??666,220,0,0q?5Z      $??20(?5[      $??Times New Roman)?5\      $??	ExtraBold"?5]      $	??Top=?5_      ??5^      ?? *SubmitCommand+5_      5b      5c      