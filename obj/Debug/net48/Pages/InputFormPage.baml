   M S B A M L   `   `   `  ����O  KWPFLateTimesheetsApp, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null;"clr-namespace:WPFLateTimesheetsAppWPFLateTimesheetsApp  5   	   (    "WPFLateTimesheetsApp.InputFormPage   5      e aSystem.Windows.Controls.Ribbon, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089\ XPresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35R NWindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35W SPresentationCore, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35F 9http://schemas.microsoft.com/winfx/2006/xaml/presentation     5   	   R NSystem.Xaml, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     5   	   @d2http://schemas.microsoft.com/expression/blend/2008     5   	   Jmc;http://schemas.openxmlformats.org/markup-compatibility/2006     5   	   4local"clr-namespace:WPFLateTimesheetsApp     5   	     N� Title$  	InputForm��5   	   $	��450��6   $	��800��6(   .��5	      �  � RowDefinitions 5
   
   �� 5      $
��Auto �6   �� 5      $
��Auto �6   �� 5      $
��Auto �6   �� 5      $
��Auto �6   �� 5      $
��Auto �6   �� 5      $
��Auto �6   �� 5      $
��Auto �6   �� 5      $
��Auto �6    � ColumnDefinitions 5   
   �� 5      $
��Auto �6   �� 5      $
��Auto �6   �� 5      $
��Auto �6   �� 5      $��* �6   .+�5   
   �� $��0��6   $��0��6   .��6-   Selected File:6>   �� 5   
   -   5        $��SelectedFileTextBlock��5   '   $��Select a File��6   $	��100��6D   $��5 5 5 5q�6S   $��0��6d   $��1��6q    �� 
IsReadOnly$
 true��6�   ${�False��6�   6�   �� 5   
   -   5        $ ��BrowseTimesheetFilesButton��5      $��20��6[   $��50��6g   $��0��6r   $��2��6   .��6�   Browse6�   �� 5   
   $��1��6   $��0��6   .��6-   .,Location of the "GTE Helpful Hint.pdf" file:6\   �� 5   
   -   5        $��AttachmentFileTextBlock��5   '   $��Select a File��6   $	��100��6F   $��5 5 5 5q�6U   $��1��6f   $��1��6s   $
 true��6�   ${�False��6�   6�   �� 5   
   -   5        $!��BrowseAttachmentFilesButton��5      $��20��6]   $��50��6i   $��1��6t   $��2��6�   .��6�   Browse6�   �� 5    
   $��2��6   $��0��6   .��6-   #!Date and Time of Report Creation:6P   �� 5!   
   -   5        $��ReportCreationTimeTextBox��5!      $��2��63   $��1��6@   6R   �� 5"   
   $��3��6   $��0��6   .��6-   !Date to Send Individual Emails:6N   (  "System.Windows.Controls.DatePicker  5#   
   -   5        $ ��DateToSendIndividualEmails��5#      $��3��67   $��1��6D   6�   �� 5$   
   -   5        $ ��IndividualEmailDateProblem��5$      $��3��6   $��2��6?   $��2��6O   $y�	Collapsed=�6c   ���#A71D00.��6�   Value must be selected6�   �� 5%   
   $��4��6   $��0��6   .��6-   Date to Send Mass Email:6G     5&   
   -   5        $��DateToSendMassEmail��5&      $��4��60   $��1��6=   6y   �� 5'   
   -	   5        $��MassEmailDateProblem��5'      $��4��6   $��2��69   $��2��6I   $y�	Collapsed=�6]   ���#A71D00.��6�   31This date must be after the individual email date6�   �� 5(   
   $��5��6   $��0��6   .��6-   " Location to Save Teams Messages:6P   �� 5)   
   -
   5        $$��TeamsMessagesFileNameTextBlock��5)   '   $��Select a File��6   $	��100��6M   $��5 5 5 5q�6\   $��5��6m   $��1��6z   $
 true��6�   ${�False��6�   6�   �� 5*   
   -   5        $+��%ChooseTeamsMessagesSaveLocationButton��5*      $��20��6n   $��50��6z   $��5��6�   $��2��6�   .��6�   Browse6�   �� 5+   
   $��6��6   $��0��6   .��6-   Name of Email Sender:6D   �� 5,   
   -   5        $��EmailSenderTextBox��5,      $��6��6,   $��1��69   6K   �� 5-   
   -   5        $��
SendEmails��5-      $��	0 10 0 20q�6#   $��7��66   $��1��6C   .��6r   Send Emails6   5.      5/      