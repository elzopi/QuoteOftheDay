   M S B A M L   `   `   `  ����H  DQuoteOftheDay, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null-clr-namespace:QuoteOftheDayQuoteOftheDay  5   	       QuoteOftheDay.MainWindow   5   	   \ XPresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35R NWindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35W SPresentationCore, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35D 9http://schemas.microsoft.com/winfx/2006/xaml/presentation    5   	   R NSystem.Xaml, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     5   	   @d2http://schemas.microsoft.com/expression/blend/2008     5   	   Jmc;http://schemas.openxmlformats.org/markup-compatibility/2006     5   	   -localclr-namespace:QuoteOftheDay     5   	   -   5          �Name$  	QODWindow��5   	    � Title$ Quote of the Day��5   	   $	��352��6"   $	��494��6/   $
��Bold"�6;    � ToolTip$ Select an option��6M    � WindowStyle$ ThreeDBorderWindow=�6h    � Icon$ Destime-pumps-transp.ico��6�   ��5	      �� 5
   
    �� Color$ 	#FFFDEDF4��6   .��5      �  � RowDefinitions 5   
   �� 5      $
��125* �6   �� 5      $
��221* �6    � ColumnDefinitions 5   
   �� 5      $��0* �6   �� 5      $	��20* �6   �� 5      $	��90* �6   �� 5      $	��35* �6   �� 5      $
��100* �6   �� 5      $
��185* �6   �� 5      $
��151* �6   .+�5   
   �� -   5         ��Name$ TextBox1��5      $
��Left=�6$   $��38��6?   	 �� TextWrapping$
	 Wrap=�6K   $	��Top=�6_   $	��285��6w   $��9,32,0,0q�6�   $
��Bold"�6�   $��Italic$�6�   $��2��6�   $��4��6�   ���	#FFE7FDFA6�   �� 5   
   -   5        $ TextBox2��5      $
��Left=�6$   $	��195��6?   $��	9,100,0,0q�6L   $
	 Wrap=�6_   $	��Top=�6s   $	��290��6�   $��2��6�   $��4��6�   ���	#FFE7FDFA$��2��6�   �� 5   
   -   5        
 ��Name$
 Button1��5      $��Go��6"   $
��Left=�6/   $��35��6J   $��120,100,0,0q�6V   $	��Top=�6k   $	��110��6�   4  .System.Windows.Automation.AutomationProperties   HelpText$ Click to get Quotes��6�   $
��Bold"�6�   $��5��6�   $��2��6�   $��2��6�   �� 5   
   -   5        $
 Button2��5      $
��Done��6"   $
��Left=�61   $��40��6L   $��120,130,0,0q�6X   $	��Top=�6m   $	��110��6�   $
��Bold"�6�   $��5��6�   $��1��6�   $��2��6�   �� 5   
   -   5         ��Name$ 	ComboBox1��5      $
��Left=�6&   $��30��6A   $��
120,35,0,0q�6M   $	��Top=�6a   $	��110��6y   $
��Bold"�6�   $��5��6�   $! Select an option to proceed��6�    �� SelectedIndex$ 0��6�    � Cursor$
 Hand��6�   $��2��6�   ��5      �� 5        �� EndPoint$	 0,1&�6&    �� 
StartPoint$	 0,0&�65   .��5!      �  � Color$ 	#FFF0F0F0��6#    � Offset$ 0Q�65   � 5"      $ 	#FF9CFBB2��6#   $ 1Q�65   5#      5%      �� 5&   
   -   5         ��Name$ Label1��5&      $��Select an Option
��6    $��5��6E   $
��Left=�6U   $��30��6p   $��	120,5,0,0q�6|   $	��Top=�6�   $	��110��6�   $( "Please select an option to proceed��6�   $��2��6�   5(      5)      