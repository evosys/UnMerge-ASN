   M S B A M L   `   `   `  ���� M  IUnMerge ASN Report, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null7 clr-namespace:UnMerge_ASN_ReportUnMerge_ASN_Report  #    UnMerge_ASN_Report.MainWindow   R NWindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35W SPresentationCore, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35\ XPresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35D 9http://schemas.microsoft.com/winfx/2006/xaml/presentation    R NSystem.Xaml, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     @d2http://schemas.microsoft.com/expression/blend/2008     Jmc;http://schemas.openxmlformats.org/markup-compatibility/2006     X TXceed.Wpf.Toolkit, Version=3.3.0.0, Culture=neutral, PublicKeyToken=3e4669d2f30244f44xctk)http://schemas.xceed.com/wpf/xaml/toolkit  ,local clr-namespace:UnMerge_ASN_Report    � 
ResizeMode$  NoResize=� � Title$ 
MainWindow��$	��380��$	��400�� � WindowStartupLocation$ CenterScreen=�.��%  Xceed.Wpf.Toolkit.BusyIndicator  -     Name$ ProgressIndicator��   IsBusy .  .����  �� Padding$ 10q�#��� ��._��� .+��  � ColumnDefinitions �� $
��2.5* ��� $��* �.+��� $��0��.+��� -    ��Name$ txtOpen�� �� 
IsReadOnly . ���#EEE$��2q��� $��1��.+��� -   	 ��Name$	 BtnOpen��$��10 0 0 0q�$��2q�$
��Open��� $
��0 13q� �� $
��2.5* ��� $��* �.+��� $��0��.+��� -   $ txtDest�� . ���#EEE$��2q��� $��1��.+��� -   $	 BtnDest��$��10 0 0 0q�$��2q�$��Destination��� $	��0 7q� �� $��* ��� $��1* �.+��� $��0��.+��� $��Options Select��$��2q�$	��155��$	��Top=�$��1q����#D1D1D1.��� $
��Left=�.+��� -   
 ��Name$
 cbUnMer��$��Un-Merge��$��5 7 0 0q�$
��Left=�$	��Top=��� $	��N/A��$��5 32 0 0q�$
��Left=�$	��Top=�{�.  �� $	��N/A��$��5 57 0 0q�$
��Left=�$	��Top=�{�.  �� $	��N/A��$��5 82 0 0q�$
��Left=�$	��Top=�{�.  �� $	��N/A��$��	5 107 0 0q�$
��Left=�$	��Top=�{�.  �� $��1��.+���    logoDrawingImage#��[  $��	10 13 0 0q�$��Center=�$��Center=�$	��110��� $
��0 30q�.+��� $��1��.+��� -   $	 BtnProc��$
��90 0q�$��30��$��2q�$��Process��