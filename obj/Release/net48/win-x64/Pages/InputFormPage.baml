   M S B A M L   `   `   `  ���� O  KWPFLateTimesheetsApp, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null;"clr-namespace:WPFLateTimesheetsAppWPFLateTimesheetsApp  (    "WPFLateTimesheetsApp.InputFormPage   \ XPresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35W SPresentationCore, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35e aSystem.Windows.Controls.Ribbon, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089R NWindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35F 9http://schemas.microsoft.com/winfx/2006/xaml/presentation     R NSystem.Xaml, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     @d2http://schemas.microsoft.com/expression/blend/2008     Jmc;http://schemas.openxmlformats.org/markup-compatibility/2006     4local"clr-namespace:WPFLateTimesheetsApp       N� Title$  	InputForm��$	��450��$	��800��.���  � RowDefinitions �� $
��Auto ��� $
��Auto ��� $
��Auto ��� $
��Auto ��� $
��Auto ��� $
��Auto ��� $
��Auto ��� $
��Auto � � ColumnDefinitions �� $
��Auto ��� $
��Auto ��� $
��Auto ��� $��* �.+��� $��0��$��0��.��Selected File:�� -   $��SelectedFileTextBlock��$��Select a File��$	��100��$��5 5 5 5q�$��0��$��1�� �� 
IsReadOnly$
 true��${�False���� -   $ ��BrowseTimesheetFilesButton��$��20��$��50��$��0��$��2��.��Browse�� $��1��$��0��.��Location of Attachment File:�� -   $��AttachmentFileTextBlock��$��Select a File��$	��100��$��5 5 5 5q�$��1��$��1��$
 true��${�False���� -   $!��BrowseAttachmentFilesButton��$��20��$��50��$��1��$��2��.��Browse�� $��2��$��0��.��#!Date and Time of Report Creation:�� -   $��ReportCreationTimeTextBox��$��2��$��1���� $��3��$��0��.��!Date to Send Individual Emails:(  "System.Windows.Controls.DatePicker  -   $ ��DateToSendIndividualEmails��$��3��$��1���� -   $ ��IndividualEmailDateProblem��$��3��$��2��$��2��$y�	Collapsed=����#A71D00.��Value must be selected�� $��4��$��0��.��Date to Send Mass Email:  -   $��DateToSendMassEmail��$��4��$��1���� -	   $��MassEmailDateProblem��$��4��$��2��$��2��$y�	Collapsed=����#A71D00.��31This date must be after the individual email date�� $��5��$��0��.��" Location to Save Teams Messages:�� -
   $$��TeamsMessagesFileNameTextBlock��$��Select a File��$	��100��$��5 5 5 5q�$��5��$��1��$
 true��${�False���� -   $+��%ChooseTeamsMessagesSaveLocationButton��$��20��$��50��$��5��$��2��.��Browse�� $��6��$��0��.��Name of Email Sender:�� -   $��EmailSenderTextBox��$��6��$��1���� -   $��
SendEmails��$��	0 10 0 20q�$��7��$��1��.��Send Emails