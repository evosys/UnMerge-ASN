   M S B A M L   `   `   `  ����M  IUnMerge ASN Report, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null7 clr-namespace:UnMerge_ASN_ReportUnMerge_ASN_Report  5   	   #    UnMerge_ASN_Report.MainWindow   5   	   R NWindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35W SPresentationCore, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35\ XPresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35D 9http://schemas.microsoft.com/winfx/2006/xaml/presentation    5   	   R NSystem.Xaml, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     5   	   @d2http://schemas.microsoft.com/expression/blend/2008     5   	   Jmc;http://schemas.openxmlformats.org/markup-compatibility/2006     5   	   X TXceed.Wpf.Toolkit, Version=3.3.0.0, Culture=neutral, PublicKeyToken=3e4669d2f30244f44xctk)http://schemas.xceed.com/wpf/xaml/toolkit  5   	   ,local clr-namespace:UnMerge_ASN_Report  5   	     � 
ResizeMode$  NoResize=�5	   	    � Title$ 
MainWindow��5
   	   $	��380��6   $	��400��6)    � WindowStartupLocation$ CenterScreen=�65   .��5      %  Xceed.Wpf.Toolkit.BusyIndicator  -   5          Name$ ProgressIndicator��5         IsBusy .  .��5   
   ��  �� Padding$ 10q�6   #��� ��6   ._�5      �� .+�5      �  � ColumnDefinitions 5      �� 5      $
��2.5* �6+   �� 5      $��* �6+   .+�5      �� $��0��6!   .+�5      �� -   5         ��Name$ txtOpen��5   "    �� 
IsReadOnly . ���#EEE$��2q�6W   5      �� 5      $��1��6!   .+�5      �� -   5        	 ��Name$	 BtnOpen��5   !   $��10 0 0 0q�6H   $��2q�6Z   $
��Open��6f   5      5      � 5      $
��0 13q�6    5       �� 5!      $
��2.5* �6+   �� 5"      $��* �6+   .+�5$      �� $��0��6!   .+�5%      �� -   5        $ txtDest��5%   "    . ���#EEE$��2q�6W   5&      �� 5'      $��1��6!   .+�5(      �� -   5        $	 BtnDest��5(   !   $��10 0 0 0q�6H   $��2q�6Z   $��Destination��6f   5)      5*      � 5.      $	��0 7q�6    5/      �� 50      $��* �6+   �� 51      $��1* �6+   .+�53      �� $��0��6!   .+�54      �� $��Options Select��6#   $��2q�6;   $	��155��6G   $	��Top=�6T   $��1q�6l   ���#D1D1D1.��55      � $
��Left=�6#   .+�56   "   �� -   5        
 ��Name$
 cbUnMer��56   +   $��Un-Merge��6<   $��5 7 0 0q�6O   $
��Left=�6`   $	��Top=�6{   �� 57   "   $	��N/A��6+   $��5 32 0 0q�69   $
��Left=�6K   $	��Top=�6f   {�.  6~   �� 58   "   $	��N/A��6+   $��5 57 0 0q�69   $
��Left=�6K   $	��Top=�6f   {�.  6~   �� 59   "   $	��N/A��6+   $��5 82 0 0q�69   $
��Left=�6K   $	��Top=�6f   {�.  6~   �� 5:   "   $	��N/A��6+   $��	5 107 0 0q�69   $
��Left=�6L   $	��Top=�6g   {�.  6   5;      5<      5=      �� 5>      $��1��6!   .+�5?      ��    logoDrawingImage#��[  6    $��	10 13 0 0q�6K   $��Center=�6^   $��Center=�6{   $	��110��6�   6�   5@      5A      � 5E      $
��0 30q�6   .+�5F      �� $��1��6!   .+�5G      �� -   5        $	 BtnProc��5G   !   $
��90 0q�6H   $��30��6V   $��2q�6b   $��Process��6n   5H      5I      5K      5M      5N      5O      