   M S B A M L   `   `   `  ����X  TComponentManagementSystem-1.0, Version=1.0.0.0, Culture=neutral, PublicKeyToken=nullO,clr-namespace:ComponentManagementSystem_1._0ComponentManagementSystem_1._0  5   	   /    )ComponentManagementSystem_1._0.MainWindow   5   	   R NWindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35W SPresentationCore, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35\ XPresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35D 9http://schemas.microsoft.com/winfx/2006/xaml/presentation    5   	   R NSystem.Xaml, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     5   	   @d2http://schemas.microsoft.com/expression/blend/2008     5   	   Jmc;http://schemas.openxmlformats.org/markup-compatibility/2006     5   	   >local,clr-namespace:ComponentManagementSystem_1._0     5   	     � Title$  
MainWindow��5   	   $	��450��6   $	��800��6)    � WindowStartupLocation$ CenterScreen=�65    � WindowState$ 	Maximized=�6Z   ���	#FFDAD7D7 � 	Resources 5	       �� MergedDictionaries �� Source�� 5
   
    5      �� 5      $> 8/ComponentManagementSystem-1.0;component/css/Styles.xaml?�6%   6i   %    5      .��5      �  � RowDefinitions 5   
   �� 5      $��3* �6   �� 5      $	��92* �6   �� 5      $��5* �6    � ColumnDefinitions 5   
   �� 5      $	��80* �6   6-   �� 5      $	��20* �6   6-   .+�5   
   ��  �� Orientation$ 
Horizontal=�6   $��0��6.   $��0��6;   $��2��6K   .+�5      �� $��12(�6   $��File Name : ��6"   6;   �� 5      -   5        	 ��Name$	 txtFileName��5      $��2q�6-   
�������6Q   �� 5      $��12(�6   $	�� | ��6"   62   �� 5      $��12(�6   $��Mode : ��6"   66   �� 5      -   5        $	 txtMode��5      $��2q�6)   
�������6N   5       �� 5!   
   $��1��6   $��0��6$   .��5"      &   System.Windows.Controls.DataGrid  -   5        
  Name$
 mainDataGrid��5"         DataGridStyle#��[  6-   6U   5#      �� 5$   
   $��1��6   $��1��6%   $��Visible=�65   .��5%      � -   5         �Name$ processGrid��5%       5&      �� 5'      $��1* �6'   �� 5(      $��1* �6'   �� 5)      $��1* �6'   5+      5,      �� 5-   
   $ 
Horizontal=�6   $��2��6/   $��0��6<   $��2��6L   .+�5.      �� -   5        $��Open Process File��5.      
��� ���$��5q�5/      $��5,5,5,5q�6    � Cursor$
 Hand��6-   6=   �� 50      -   5        $��Open Data File��50      
��� ���$��5q�6e   $��5,5,5,5q�6q   $
 Hand��6�   6�   �� 51      -   5        $��Generate��51      ���	#FFA0FCA0$��5,5,5,5q�6_   $��5q�6q   $
 Hand��6}   �� 52      -   5        $
��Save��52      ���	#FFA0FCA0$��5,5,5,5q�6R   $��5q�6d   $
 Hand��6p   �� 53      -	   5        $��Add Mode��53      ���	#FFDFDFDF$��5,5,5,5q�6Y   $��5q�6k   $
 Hand��6w   �� 54      -
   5        $��	Edit Mode��54      ���	#FFDFDFDF$��5,5,5,5q�6]   $��5q�6o   $
 Hand��6{   �� 55      -   5        $��Delete Mode��55      ���	#FFDFDFDF$��5,5,5,5q�6_   $��5q�6q   $
 Hand��6}   �� 56      -   5        $	 txtError��56       $�� ��6   
���  ��$��5,5,5,10q�6C   6W   �� 57      -   5        $	 
txtSuccess��57       $�� ��6   
��� � �$��5,5,5,10q�6G   6[   58      5:      5<      