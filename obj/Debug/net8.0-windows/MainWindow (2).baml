   M S B A M L   `   `   `  ����  
Morden App    Morden_App.MainWindow   5   	   _ [MaterialDesignThemes.Wpf, Version=5.2.1.0, Culture=neutral, PublicKeyToken=df2a72020bd7962aFmaterialDesign1http://materialdesigninxaml.net/winfx/xaml/themes  5   	   W SPresentationCore, Version=8.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35\ XPresentationFramework, Version=8.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35e aSystem.Windows.Controls.Ribbon, Version=8.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089R NWindowsBase, Version=8.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35F 9http://schemas.microsoft.com/winfx/2006/xaml/presentation     5   	   R NSystem.Xaml, Version=8.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     5   	     � Title$  
MainWindow��5   	   $	��450��6   $	��800��6)   
�������.��5      ]�  � Cursor$  ��6   .+�5	   
   �� $
��Left=�6   $	��200��6+   ���#2E3B4E.+�5
      �� $��🌟 Test App��6   $��18(�6+   
�������$��20q�6L   �� 5      -   5        $��🏠 首頁��5         MaterialDesignFlatButton#��[  5      
�������$��5q�5      5      �� 5      -   5         ��Name$ ExportSideButton��5      $��📤 匯出 Excel��5      #��[  5      
�������$��5q�5      5      �� 5      $��📑 查看錯誤報告��6   #��[  6)   
�������$��5q�6n   �� 5      -   5        $��👁️‍🗨️ 預覽��5      #��[  5      
�������$��5q�5      5      �� 5      $��⚙️ 設定��6   #��[  6%   
�������$��5q�6j   5      � 5   
   -   5         �Name$ 	MainPanel��5      .+�5!      �� $	��Top=�5"   	   $��Center=�5#   	   $��0,30,0,0q�5$   	   $��24(�5%   	   ���#444�� 5'   
    � Language$ zh-tw�6   $�Welcome��6   5(      �� 5+      -   5         ��Name$ UserComboBox��5,   	   $	��172��5-   	   $��42��5.   	   $	��Top=�5/   	   $
��Left=�50   	   $��
396,25,0,0q�51   	      MaterialDesignFilledComboBox#��[ 52   	   )  #MaterialDesignThemes.Wpf.HintAssist   Hint$ 選擇使用者��53   	    �� SelectedIndex$ 0��54   	   .@�55      �� $��Jason��6   �� 56      $
��Jane��6   �� 57      $	��Wei��6   �� 58      $��JDLee��6   �� 59      $��Peggy��6   �� 5:      $
��Fish��6   5;      �� 5>      -   5         ��Name$ DateTextBlock��5?   	   $	��Top=�5@   	   $��Center=�5A   	   $��0,84,0,0q�5B   	   $��16(�5C   	   ���#888$��25��5E   &   ��6	   �� *CurrentDate+6&   �� 5H      -   5        	 ��Name$	 FilePathTextBox��5I   	   $	��433��5J   	   $��50��5K   	   $	��Top=�5L   	   $
��Left=�5M   	   $ 檔案路徑��5N   	    % !MaterialDesignFloatingHintTextBox#��[ 5O   	   $��
10,100,0,0q�5P   	   �� 5S      -   5        $��選擇檔案��5T   	   $	��120��5U   	   $��36��5V   	   $	��Top=�5W   	   $
��Left=�5X   	   $��448,114,0,0q�5Y   	     MaterialDesignRaisedButton#��[ 5Z   	   $
 Hand��5[   "   5\      � 5_   
   -	   5        $ PreviewPanel��5_      $y�	Collapsed=�6%   .+�5a      �� 
 �� Orientation$
 
Horizontal=�6   $��Right=�5b      $	��Top=�5c      $��
10,10,10,0q�5d      .+�5e      �� -
   5        $ ExportPreviewButton��5f      $��📤 匯出 Excel��5e      #��[ 5g      ���#4CAF50
�������$	��5,0q�5j      5k      �� 5l      -   5        $��🔙 返回��5l         MaterialDesignOutlinedButton#��[ 5m      $	��5,0q�5n      5o      5p      �� 5r      $
��Auto=�6   $
��Auto=�5s      $��0,50,0,0q�5t      .��5u      &   System.Windows.Controls.DataGrid  -   5          Name$ PreviewGrid��5u         AutoGenerateColumns$ False��5v      $��10q�5w         CanUserAddRows$ False��5x         Columns 5y      0  *System.Windows.Controls.DataGridTextColumn  5z      ,  &System.Windows.Controls.DataGridColumn   Header$ 項次��6-      Width5  /System.Windows.Controls.DataGridLengthConverter$ 50 6Z   1  +System.Windows.Controls.DataGridBoundColumn   Binding 69   �� *IndexDisplay+6Z     5{      $ 名稱��6-   $	 100 6V    69   �� *
PartName+6V     5|      $ 規格��6-   $	 150 6R    69   �� *Spec+6R     5}      $ 數量��6-   $ 70 6]    69   �� *QuantityDisplay+6]     5~      $ 代號��6-   $	 100 6R    69   �� *Code+6R     5      $ 單價��6-   $ 70 6^    69   �� *UnitPriceDisplay+6^     5�      $ 小計��6-   $ 80 6]    69   �� *SubtotalDisplay+6]     5�      $ 廠商��6-   $	 150 6T    69   �� *Vendor+6T     5�      $ 中國廠商��6-   $	 150 6X    6;   �� *
VendorCN+6X     5�      $ 備註��6-   $	 200 6R    69   �� *Note+6R   5�      5�      5�      5�      5�      