   M S B A M L   `   `   `  ����  GUIWISPDataHandling$    GUIWISPDataHandling.MainWindow   5      W SPresentationCore, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35R NWindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35\ XPresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35D 9http://schemas.microsoft.com/winfx/2006/xaml/presentation    R NSystem.Xaml, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     5      H DMahApps.Metro, Version=1.6.5.1, Culture=neutral, PublicKeyToken=null;Controls,http://metro.mahapps.com/winfx/xaml/controls  5      R NMahApps.Metro.IconPacks, Version=2.3.0.4, Culture=neutral, PublicKeyToken=null=	iconPacks-http://metro.mahapps.com/winfx/xaml/iconpacks  5      :d2http://schemas.microsoft.com/expression/blend/2008  5      Dmc;http://schemas.openxmlformats.org/markup-compatibility/2006  6U      AccentColorBrush(  "MahApps.Metro.Controls.MetroWindow    	GlowBrush#  �   5       � Title$ 	QPMS WISP��5       � WindowStartupLocation$ CenterScreen=�5	      $	��964��6P      RightWindowCommands 5      +  %MahApps.Metro.Controls.WindowCommands  5   
   .@�5      �� -   5        .��5      ��  �� Orientation$ 
Horizontal=�6   .+�5      ,  &MahApps.Metro.IconPacks.PackIconModern  $��24��6/   $��24��6:   E AControlzEx, Version=3.0.2.4, Culture=neutral, PublicKeyToken=null�  �ControlzEx.PackIconBase`1[[MahApps.Metro.IconPacks.PackIconModernKind, MahApps.Metro.IconPacks, Version=2.3.0.4, Culture=neutral, PublicKeyToken=null]]   Kind$ Acorn=�6F   �� 5      $��4 0 0 0q�6    $��Center=�61   $
��Help��6L   5      5      5      .��5      � .+�5   
   �� $��Center=�6   $	��Top=�63   $��288,10,256,0q�6K   $	��165��6a   $	��412��6n   .+�5      �� $	��120��6   $ ��/Assets/quinovicgllogo.jpg��6    $��0,0,0,0q�6D   $��Center=�6U   �� 5       $)��#QPMS Water Invoice System Processor��6   $��
0,-20,20,0q�6B   $��20(�6V   $
��Bold"�6d   $��Arial Rounded MT Bold)�6v   5"      �� 5%   
   -   5         ��Name$ dtGrid1��5%      $��Center=�6&   $��
10,230,0,0q�6C   .+�5&      &   System.Windows.Controls.DataGrid  -   5          Name$ dtGrid��5&   a      AutoGenerateColumns . 	   
IsReadOnly	 . $��Center=�6D   $��Center=�6q   $y�Hidden=�6�   ��. $
��Auto=�5'      $
��Auto=�5(      $	��300��69   5)      �� 50   
   $��Bottom=�6   $��0,0,0,60q�60   �� 51   
   $��0��6   $ 
Horizontal=�6"   $��Bottom=�6;   $
��Left=�6V   $��0,0,0,15q�6q   .+�55      �� -   5        
 ��Name$
 btnAddInvoice��55   ,   $��Add Invoices��6   $��15(�6C   $
��Bold"�6Q   
�������$��45��6�   $��0,0,2,0q�6�   �� 56      -   5        $
 btnOpen��56   -   $��Generate Data��6   $��15(�6>   $
��Bold"�6L   
�������$��45��6�   $��0,0,2,0q�6�   �� 57      -   5        $
 btnGoogleSheet��57   9   $��Inspect/Modify Batch Data��6   $��15(�6Q   $
��Bold"�6_   
�������$��45��6�   $��0,0,2,0q�6�   $,��&Segoe UI, Lucida Sans Unicode, Verdana)�6�   �� 58      -   5        $
 
btnProcess��58   -   $��Process Batch��6   $��15(�6A   $
��Bold"�6O   
���2͚�$��45��6�   $��0,0,2,0q�6�   $y�Hidden=�6�   5;      5?      5@      