   M S B A M L   `   `   `  ���� H  DQuoteOftheDay, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null-clr-namespace:QuoteOftheDayQuoteOftheDay      QuoteOftheDay.MainWindow   \ XPresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35R NWindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35W SPresentationCore, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35D 9http://schemas.microsoft.com/winfx/2006/xaml/presentation    R NSystem.Xaml, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     @d2http://schemas.microsoft.com/expression/blend/2008     Jmc;http://schemas.openxmlformats.org/markup-compatibility/2006     -localclr-namespace:QuoteOftheDay     -     �Name$  	QODWindow�� � Title$ Quote of the Day��$	��352��$	��494��$
��Bold"� � ToolTip$ Select an option�� � WindowStyle$ ThreeDBorderWindow=� � Icon$ Destime-pumps-transp.ico������  �� Color$ 	#FFFDEDF4��.���  � RowDefinitions �� $
��125* ��� $
��221* � � ColumnDefinitions �� $��0* ��� $	��20* ��� $	��90* ��� $	��35* ��� $
��100* ��� $
��185* ��� $
��151* �.+��� -    ��Name$ TextBox1��$
��Left=�$��38��	 �� TextWrapping$
	 Wrap=�$	��Top=�$	��285��$��9,32,0,0q�$
��Bold"�$��Italic$�$��2��$��4�����	#FFE7FDFA�� -   $ TextBox2��$
��Left=�$	��195��$��	9,100,0,0q�$
	 Wrap=�$	��Top=�$	��290��$��2��$��4�����	#FFE7FDFA$��2���� -   
 ��Name$
 Button1��$��Go��$
��Left=�$��35��$��120,100,0,0q�$	��Top=�$	��110��4  .System.Windows.Automation.AutomationProperties   HelpText$ Click to get Quotes��$
��Bold"�$��5��$��2��$��2��{�.  $y�Hidden=��� -   $
 Button2��$
��Done��$
��Left=�$��40��$��120,139,0,0q�$	��Top=�$	��110��$
��Bold"�$��5��$��1��$��2���� -    ��Name$ 	ComboBox1��$
��Left=�$��30��$��
120,35,0,0q�$	��Top=�$	��110��$
��Bold"�$��5��$! Select an option to proceed�� �� SelectedIndex$ 0�� � Cursor$
 Hand��$��2������  �� EndPoint$	 0,1&� �� 
StartPoint$	 0,0&�.���  � Color$ 	#FFF0F0F0�� � Offset$ 0Q�� $ 	#FF9CFBB2��$ 1Q��� -    ��Name$ Label1��$��Select an Option
��$��5��$
��Left=�$��30��$��	120,5,0,0q�$	��Top=�$	��110��$( "Please select an option to proceed��$��2��