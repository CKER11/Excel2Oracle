   M S B A M L   `   `   `  ���� G  CExcel2Oracle, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null+clr-namespace:Excel2OracleExcel2Oracle      Excel2Oracle.MainWindow   \ XPresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35R NWindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35W SPresentationCore, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35D 9http://schemas.microsoft.com/winfx/2006/xaml/presentation    R NSystem.Xaml, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     @d2http://schemas.microsoft.com/expression/blend/2008     Jmc;http://schemas.openxmlformats.org/markup-compatibility/2006     ,localclr-namespace:Excel2Oracle       � Title$  
MainWindow��$	��500��$
��1000�� � Icon$ Images/excel2oracle.png��.��� -    �Name$ 	wholeGrid��$��0q�$��Stretch=�$��Stretch=� � ColumnDefinitions �� $	��100��$	��35* ��� $	��390��$	��28* ��� $	��200��$	��33* �.+��� $��5��$��0���� $��5��$��1��R� -    R�Name$ TemplateTreeView��$��Stretch=�$��Stretch=�$��0��$��0,0,10,0q�$��0q�����  �� 
TargetType" Q�.���  �� Property �@Q�
IsExpanded �� Value �� *
IsExpanded+����  �� ItemsSource �� *
Children+.R��� 	 �� Orientation$	 
Horizontal=�.+��� $��Center=�$��16��$��16��$��0,0,2,2q����� *Icon+�� $��nodeName��$��Center=����� *DisplayName+�� -   $��Center=�$��Center=�$��15��$��15��$��20,0,0,0q�y��� *	Visiblity+���� 
 �� ImageSource$
 Images/refesh.jpg�� j� Stretch$
 Fill=� � ToolTip �� $��Center=�$
��Wrap=�$	��200������ *Path+� -   $ opGrid��$��1�� � RowDefinitions �� $��40 ��� $��40 ��� $��40 ��� $	��100 ���  �� $	��90* ��� $��40 �.+��� -    ��Name$ lblFileSelect��$��当前文件：��$
��Left=�$��Bottom=�$��2��$��0,0,0,7q��� -    ��Name$ txtFilePath��$
��Left=�$��24�� �� TextWrapping$ NoWrap=�$�� ��$��Center=�$��2��$��0��$��75,8,0,9q�$��1q�$	��168��.�� � ContextMenu �� .@�m� -   $��openImportFile��$��打开导入文件��m� -   $��openImportPath��$��打开导入文件目录��.���� -	    ��Name$ btnSelectFile��$	��...��$
��Left=�$��Center=�$��2��$��
248,10,0,9q�$��0q�$��25���� -
   $ lblIntrduction��$��当前模板：��$
��Left=�$��0,10,0,0q�$	��Top=�$��0���� -   $ lblCurTemplate��$��	未选择��$
��Left=�$��	78,11,0,0q�
���  ��$	��Top=�$��0�� �� .@�m� -   $��openTemplateFile��$��打开模板文件��m� -   $��openTemplatePath��$��打开模板文件目录���� -   $ lblIntrduction_Copy��$��当前规则：��$
��Left=�$��0,10,0,0q�$	��Top=�$��1���� -   $ 
lblCurRule��$��	未解析��$
��Left=�$��	78,11,0,0q�
���  ��$	��Top=�$��1���� -   $ btnAnalysisView��$��查看规则��$��Right=�$��1��$��0��$��Center=�$��70��$��	0,12,15,6q�$��2���� -   $ btnAnalysis��$��解析规则��$��Right=�$��0��$��0��$��Center=�$��70��$��	0,12,15,6q�$��2���� -   $ 
btnGetDate��$��提取数据��$��Right=�$��2��$��0��$��Center=�$��70��$��	0,10,15,8q�$��2���� -   $ 
txtProcess��$��3��$��Stretch=�$
 Wrap=�$-��'此处显示解析和提取数据进度��$��Stretch=�$��2��" �� VerticalScrollBarVisibility$ Visible=�$��	0,10,15,0q��� -   $��
viewScroll��$��4��$��0��$��2��$��Stretch=�$��Stretch=�$��	0,10,15,0q�.���� -   $��spView��$��Stretch=�$��Stretch=�� $��2��$��Stretch=�$��Stretch=� �� $��30 ��� $��40 ��� $��50 ��� .+��� -   $ label��$��关系库连接字符串:��$
��Left=�$	��Top=�$��0��$��10,4,0,0q��� -   $ txtDBConnectInfo��$��1��$��Stretch=�$
 Wrap=�$<��6data source=myserver;user id=lpf;password=Lpf19941003;��$	��Top=�$��	10,0,12,0q�$��40���� -   $ btnTestDBConnect��$��测试��$��2��$��Right=�$��Bottom=�$��49��$��0,0,76,8q��� -   $ 	btnImport��$��导入��$��2��$��Right=�$��Bottom=�$��48��$��0,0,12,8q��� -   $ txtImportProcess��$��3��$��Stretch=�$
 Wrap=�$��此处显示导入进度��$��Stretch=�$��
10,10,10,0q��� -   $ btnSaveConnect��$��保存��$��2��$��Right=�$��Bottom=�$��49��$��	0,0,140,8q�