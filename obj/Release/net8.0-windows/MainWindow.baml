   M S B A M L   `   `   `  ����   SchBom convert    SchBom_Convert.MainWindow   _ [MaterialDesignThemes.Wpf, Version=5.2.1.0, Culture=neutral, PublicKeyToken=df2a72020bd7962aFmaterialDesign1http://materialdesigninxaml.net/winfx/xaml/themes  W SPresentationCore, Version=8.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35\ XPresentationFramework, Version=8.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35e aSystem.Windows.Controls.Ribbon, Version=8.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089R NWindowsBase, Version=8.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35F 9http://schemas.microsoft.com/winfx/2006/xaml/presentation     R NSystem.Xaml, Version=8.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml       � Title$  SchBom Format Convert��$	��450��$	��800��
������� � Cursor$  �� � 
ResizeMode$ NoResize=� � WindowStartupLocation$ CenterScreen=�.��]� $  ��.+��� $
��Left=�$	��200�����#2E3B4E$  ��.+��� $��🌟 SchBom App��$��18(�
�������$��20q��� -   $��🏠 首頁��   MaterialDesignFlatButton#��[  
�������$��5q��� -    ��Name$ ExportSideButton��$��📤 匯出 Excel��#��[  
�������$��5q��� -   $��📝 預覽��#��[  
�������$��5q��� -   $��⚙️ 人員設定��#��[  
�������$��5q�� -    �Name$ 	MainPanel��.+��� $	��Top=�$��Center=�$��0,30,0,0q�$��24(����#444��  � Language$ zh-tw�$�Welcome���� -    ��Name$ DateTextBlock��$	��Top=�$��Center=�$��0,84,0,0q�$��16(����#888$��25������ *CurrentDate+�� -    ��Name$ FilePathTextBox��$	��433��$��50��$	��Top=�$
��Left=�)  #MaterialDesignThemes.Wpf.HintAssist	   Hint$	 檔案路徑�� % !MaterialDesignFloatingHintTextBox#��[ $��
10,100,0,0q�$  ���� -   $��選擇檔案��$	��120��$��36��$	��Top=�$
��Left=�$��448,114,0,0q�  MaterialDesignRaisedButton#��[ $
 Hand���� -	   
 ��Name$
 AutoOpenFileCheckBox��$&�� 📂 匯出後自動開啟檔案��$	��Top=�$
��Left=�$��
10,165,0,0q�$��14(����#666 d� 	IsChecked$ FalseY�� -
   $ SettingsPanel��$y�	Collapsed=�.+��� $
��Auto=�$
��Auto=�.��� $��10q�$	��650�� � ColumnDefinitions �� $��* �$	��390���� $
��Auto �$��25���� $��* �$��0��.+��� $��0��.+��� $0��*👨‍💼 開發人員設定 👩‍💼��$��24(�$
��Bold"����#2E3B4E$��0,0,0,20q�$��Center=��� $��研發助理人員��$��16(�$
��Bold"�$��0,0,0,10q����#444� $��0,0,0,20q� �� $��* ��� $
��Auto ��� $
��Auto �.+��� -    ��Name$ RDAssistantComboBox��$��0��$��42��$��Stretch=�   MaterialDesignFilledComboBox#��[ $'	 !選擇或輸入研發助理人員�� �� 
IsEditable$
 True�� �� SelectedIndex$ 0�� �� Text �� *SelectedRDAssistant+ �� Mode$ TwoWay=� �� UpdateSourceTrigger$ PropertyChanged=�.@��� $��Peggy���� $
��Fish���� -   $��1�� � ToolTip$ 新增人員��$��40��$��36��$��5,0,0,0q�#��[ ���#4CAF50
�������$��0q�.��'  !MaterialDesignThemes.Wpf.PackIcon     Kind$ PlusBoxMultipleOutline=�$��18��$��18���� -   $��2��$ 刪除人員��$��40��$��36��$��5,0,0,0q�#��[ 
����p��
�������$��0q�.��  $ DeleteEmptyOutline=�$��18��$��18���� $��Layout人員��$��16(�$
��Bold"�$��0,0,0,10q����#444� $��0,0,0,20q� �� $��* ��� $
��Auto ��� $
��Auto �.+��� -   $ LayoutPersonComboBox��$��0��$��42��$��Stretch=�#��[ $!	 選擇或輸入Layout人員��$
 True��$ 0�� �� *SelectedLayoutPerson+$ TwoWay=�$ PropertyChanged=�.@��� $��未定���� $	��WEI���� $
��Jane���� $
��Wuct���� $��JDLee���� $��Jason���� -   $��1��$ 新增人員��$��40��$��36��$��5,0,0,0q�#��[ ���#4CAF50
�������$��0q�.��  $ PlusBoxMultipleOutline=�$��18��$��18���� -   $��2��$ 刪除人員��$��40��$��36��$��5,0,0,0q�#��[ 
����p��
�������$��0q�.��  $ DeleteEmptyOutline=�$��18��$��18���� $��線路設計人員��$��16(�$
��Bold"�$��0,0,0,10q����#444� $��0,0,0,20q� �� $��* ��� $
��Auto ��� $
��Auto �.+��� -   $ CircuitDesignerComboBox��$��0��$��42��$��Stretch=�#��[ $'	 !選擇或輸入線路設計人員��$
 True��$ 0�� �� *SelectedCircuitDesigner+$ TwoWay=�$ PropertyChanged=�.@��� $��未定���� $	��LSP���� $
��Jane���� $��Jason���� $��Kevin���� $��Yanchi���� -   $��1��$ 新增人員��$��40��$��36��$��5,0,0,0q�#��[ ���#4CAF50
�������$��0q�.��  $ PlusBoxMultipleOutline=�$��18��$��18���� -   $��2��$ 刪除人員��$��40��$��36��$��5,0,0,0q�#��[ 
����p��
�������$��0q�.��  $ DeleteEmptyOutline=�$��18��$��18����  �� Orientation$ 
Horizontal=�$��Center=�$��
-70,10,0,0q�.+��� -   $��💾 儲存設定��#��[ ���#4CAF50
�������$	��120��$��36��$
��70,0q��� -   $��🔙 返回��   MaterialDesignOutlinedButton#��[ $	��120��$��36��$	��0,0q��� $��2�����#E0E0E0$��2q����#F8F9FA �� CornerRadius$ 10��$��0,0,0,0q� �� Padding$ 20q�$��Center=�$��Center=�._��� .+��� $��📋 當前設定��$��14(�$
��Bold"����#666$��0,0,0,10q��� -   $ CurrentRDAssistantText��$��研發助理：Peggy��$��12(����#888$
��0,10q��� -   $ CurrentLayoutPersonText��$��Layout人員：未定��$��12(����#888$
��0,10q��� -   $  CurrentCircuitDesignerText��$��線路設計：未定��$��12(����#888$
��0,10q�� -   $ PreviewPanel��$y�	Collapsed=�.+��� $����📢 請再次確認替代料 (廠商/代理商)，如異常
點擊儲存格並修改替代料格式同時提供(廠商/代理商)��$
��Wrap=�$��14(�$
��Bold"�
����p��$��	10,10,0,0q��� $ 
Horizontal=�$��Right=�$	��Top=�$��
10,10,10,0q�.+��� -   $ ExportPreviewButton��$��📤 匯出 Excel��#��[ ���#4CAF50
�������$	��5,0q��� -   $��🔙 返回��#��[ $	��5,0q��� $
��Auto=�$
��Auto=�$��0,50,0,0q�.��&   System.Windows.Controls.DataGrid  -     Name$ PreviewGrid��   AutoGenerateColumns$ False��$��10q�   CanUserAddRows$ False��   Columns 0  *System.Windows.Controls.DataGridTextColumn  ,  &System.Windows.Controls.DataGridColumn   Header$ 項次��   Width5  /System.Windows.Controls.DataGridLengthConverter$ 50 1  +System.Windows.Controls.DataGridBoundColumn   Binding �� *IndexDisplay+  $ 名稱��$	 100  �� *
PartName+  $ 規格��$	 150  �� *Spec+  $ 數量��$ 70  �� *QuantityDisplay+  $ 零件編號��$	 100  �� *Code+  $ 單價��$ 70  �� *UnitPriceDisplay+  $ 小計��$ 80  �� *SubtotalDisplay+  $ 廠商��$	 150  �� *Vendor+  $ 中國廠商��$	 150  �� *
VendorCN+  $ 備註��$	 200  �� *Note+