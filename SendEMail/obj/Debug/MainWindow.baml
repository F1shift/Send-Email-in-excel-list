   M S B A M L   `   `   `  ����D  @SendEMail, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null;"clr-namespace:SendEMail.PropertiesSendEMail.Properties  5   	       SendEMail.MainWindow   5   	   R NWindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35W SPresentationCore, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35\ XPresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35D 9http://schemas.microsoft.com/winfx/2006/xaml/presentation    5   	   R NSystem.Xaml, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     5   	   6propers"clr-namespace:SendEMail.Properties     5   	   -   5          � Title$  
MainWindow��5   	   $	��335��6   $	��600��6)   ���	#FF444444$��24(�6e    � 	Resources 5       �� MergedDictionaries �� Source�� 5   
    5      �� 5	      $ Theme/Styles.xaml?�6%   6B   %    5      .��5      �  � ColumnDefinitions 5   
   �� 5      $
��Auto �6   �� 5      $��1* �6    � RowDefinitions 5   
   �� 5      $
��Auto �6   �� 5      $
��Auto �6   �� 5      $
��Auto �6   �� 5      $
��Auto �6   �� 5      $
��Auto �6   �� 5      $
��Auto �6   .+�5   
   �� $��0��6   $��0��6!   $��Center=�61   $��10q�6L   Server6`   �� 5   
   -   5         ��Name$ 	tb_Server��5      $��0��6%   $��1��62   $��Center=�6�   ��6B   �� #  SendEMail.Properties.Settings   Default �� Source# Z 	 �� Path$	 
smtpServer�.��6�   5      �� 5   
   $��1��6   $��0��6!   $��Center=�61   $��10q�6L   Port6^   �� 5   
   -   5        $ tb_Port��5      $��1��6$   $��1��61   $��Center=�6�   ��6A   �� # Z $	 smtpPort�.��6�   5      �� 5    
   $��2��6   $��0��6!   $��Center=�61   $��10q�6L   	Account6a   �� 5!   
   -   5        $ 
tb_Account��5!      $��2��6'   $��1��64   $��Center=�6�   ��6D   �� # Z $	 mailAccount�.��6�   5"      �� 5#   
   $��3��6   $��0��6!   $��Center=�61   $��10q�6L   
Password6b   G� 5$   
   -   5        
 G�Name$
 tb_Pwd��5$      $��3��6'   $��1��64   $��Center=�6D   5%      �� 5&   
   $��4��6   $��0��6!   $��Center=�61   $��10q�6L   
Excel File6d   �� 5'   
   $��4��6   $��1��6#    �� Orientation$ 
Horizontal=�63   .+�5(      �� -   5        $��10q�5(      .��69   Choose File6F   �� 5)      -   5         ��Name$ tb_ExcelFilePath��5)      $��Center=�62   $��10,0,0,0q�6M   6a   5*      �� 5+   
   -   5        $��5��5+      $��0��6   $��2��6.   $��10,10,10,10q�6B   .��6n   Send E-mail6{   5,      5-      