   M S B A M L   `   `   `  ���� H  DCountWodayWpf, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null-clr-namespace:CountWodayWpfCountWodayWpf      CountWodayWpf.MainWindow   W SPresentationCore, Version=5.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35\ XPresentationFramework, Version=5.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35e aSystem.Windows.Controls.Ribbon, Version=5.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089R NWindowsBase, Version=5.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35F 9http://schemas.microsoft.com/winfx/2006/xaml/presentation     R NSystem.Xaml, Version=5.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     @d2http://schemas.microsoft.com/expression/blend/2008     Jmc;http://schemas.openxmlformats.org/markup-compatibility/2006     -localclr-namespace:CountWodayWpf     T PMahApps.Metro, Version=2.0.0.0, Culture=neutral, PublicKeyToken=51482d6f650b2b3f;controls,http://metro.mahapps.com/winfx/xaml/controls    � Title$  MDat.CountWorkDays��$	��799��$	��800��(  "MahApps.Metro.Controls.MetroWindow   	GlowBrush
 ����� � WindowStartupLocation$ CenterScreen=�.��� $��20q� � RowDefinitions �� $
��Auto ��� $
��Auto ��� $��* ��� $
��Auto �.+���  �� Orientation$ 
Horizontal=�$	��Top=�.+��� $��Chọn file Excel input:��$��Center=�$��0,0,10,0q��� -    ��Name$ InputFilePathTextBox��$	��400�� �� 
IsReadOnly$
 True��$��0,0,10,0q��� -   $��Chọn file...���� $ 
Horizontal=�$	��Top=�$��0,40,0,0q�$��1��.+��� $��Chọn file output:��$��Center=�$��0,0,10,0q��� -   $ OutputFilePathTextBox��$	��400��$
 True��$��0,0,10,0q��� -   $��Chọn nơi lưu...���� $��Debug��$��2��$��0,20,0,0q�.���� -   $ DebugTextBox��" �� VerticalScrollBarVisibility$
 Auto=�$ �� HorizontalScrollBarVisibility$
 Auto=�	 �� AcceptsReturn$
	 True��$
 True��
 �� TextWrapping$

 Wrap=��� $��Cài đặt thông tin��$��3��$��0,42,0,0q�.���� .+��� $ 
Horizontal=�$��10q�.+��� $��Ô chứa tên:��$��Center=�$��0,0,10,0q��� -   $ CellNameTextBox��$��80��$��0,0,10,0q��� $��Ô bắt đầu:��$��Center=�$��0,0,10,0q��� -   $ CellStartTextBox��$��80��$��0,0,10,0q��� $��Ô kết thúc:��$��Center=�$��0,0,10,0q��� -   $ CellEndTextBox��$��80��$��0,0,10,0q��� -	   $��Thêm dòng��$	��100���� -
   $��
Xóa dòng��$	��100��$��10,0,0,0q��� 
�������$��1q� �� CornerRadius$ 4��$��
10,10,10,0q�._��� -    ��Name$ AddedRowsListBox��$	��100���� $ 
Horizontal=�$��	10,10,0,0q�.+��� $��Key String:��$��Center=�$��0,0,10,0q��� -   $ KeyStringTextBox��$	��200���� $ 
Horizontal=�$��	10,10,0,0q�.+��� $��Giờ vào làm:��$��Center=�$��0,0,10,0q��� -   $ StartTimeTextBox��$��80��$��0,0,20,0q�$��08:30���� $��Giờ tan làm:��$��Center=�$��0,0,10,0q��� -   $ EndTimeTextBox��$��80��$��17:30���� $ 
Horizontal=�$��	10,10,0,0q�.+��� $ 
Horizontal=�$��10q�.+��� $��Tiền Phạt:��$��Center=�$��0,0,10,0q��� -   $ 
FeeTextBox��$��80��$��0,0,10,0q��� -   $��Thêm dòng��$	��100���� -   $��
Xóa dòng��$	��100��$��10,0,0,0q��� 
�������$��1q�$ 4��$��
10,10,10,0q�._��� -   $ AddedRowsFeeListBox��$	��100���� $��	10,10,0,0q�.+��� -    ��Name$ LoadingText��$��Đang xử lý...��$y�	Collapsed=�
���  ��$
��0,10q�� -    �Name$ ProgressBar�� � IsIndeterminate$
 True��$y�	Collapsed=�$��20���� -   $��Export File Chấm công��$��3��$��	557,5,0,0q�$��30��$	��Top=�