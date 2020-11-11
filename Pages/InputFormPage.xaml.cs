using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Microsoft.Win32;
using System.Diagnostics;
using WPFLateTimesheetsApp.LateTimesheetsRepositories;
using WPFLateTimesheetsApp.Controllers;
using WPFLateTimesheetsApp.EmailSenders;

namespace WPFLateTimesheetsApp
{
    public partial class InputFormPage : Page
    {
        public InputFormPage()
        {
            InitializeComponent();
        }

        public void DisplayTimesheetSelectionWindow(object sender, RoutedEventArgs e)
        {
            OpenFileDialog fileSelectionDialog = new Microsoft.Win32.OpenFileDialog();

            Nullable<bool> fileSelected = fileSelectionDialog.ShowDialog();

            if(fileSelected.HasValue && fileSelected.Value)
            { 
                this.SelectedFileTextBlock.Text = fileSelectionDialog.FileName;
            }
        }

        public void DisplayAttachmentSelectionWindow(object sender, RoutedEventArgs e)
        {
            OpenFileDialog fileSelectionDialog = new Microsoft.Win32.OpenFileDialog();

            Nullable<bool> fileSelected = fileSelectionDialog.ShowDialog();

            if(fileSelected.HasValue && fileSelected.Value)
            { 
                this.AttachmentFileTextBlock.Text = fileSelectionDialog.FileName;
            }
        }

        public void DisplayTeamsMessagesSaveSelectionWindow(object sender, RoutedEventArgs e)
        {
            SaveFileDialog fileSelectionDialog = new Microsoft.Win32.SaveFileDialog();
            fileSelectionDialog.Filter = "txt files (*.txt)|*.txt";
            fileSelectionDialog.FilterIndex = 1;

            Nullable<bool> fileSelected = fileSelectionDialog.ShowDialog();

            if(fileSelected.HasValue && fileSelected.Value)
            { 
                this.TeamsMessagesFileNameTextBlock.Text = fileSelectionDialog.FileName;
            }
        }

        public void EnsureDatesAreValid(object sender, RoutedEventArgs e)
        {
            if(!this.DateToSendIndividualEmails.SelectedDate.HasValue)
            {
                this.IndividualEmailDateProblem.Visibility = Visibility.Visible;
            }
            else
            {
                if(!this.DateToSendMassEmail.SelectedDate.HasValue || 
                    this.DateToSendIndividualEmails.SelectedDate.Value >= this.DateToSendMassEmail.SelectedDate.Value)
                {
                    this.MassEmailDateProblem.Visibility = Visibility.Visible;
                }
                else
                {
                    this.MassEmailDateProblem.Visibility = Visibility.Collapsed;
                    this.IndividualEmailDateProblem.Visibility = Visibility.Collapsed;
                }
            }
        }

        public void SendEmailsAndAlertUser(object sender, RoutedEventArgs e)
        {
            bool allFieldsValid = true;
            string timeOfReportCreation = this.ReportCreationTimeTextBox.Text; 
            allFieldsValid = (allFieldsValid && !String.IsNullOrWhiteSpace(timeOfReportCreation));
            string timeSheetFile = this.SelectedFileTextBlock.Text;
            allFieldsValid = (allFieldsValid && !String.IsNullOrWhiteSpace(timeSheetFile));
            ILateTimesheetRepository timesheetRepository = null;
            if(System.IO.Path.GetExtension(timeSheetFile) == ".txt")
            {
                timesheetRepository = new LateTimesheetsTsvRepository(timeSheetFile);
            }
            if(timesheetRepository == null)
            {
                allFieldsValid = false;
            }
            string attachmentFile = this.AttachmentFileTextBlock.Text;
            allFieldsValid = (allFieldsValid && !String.IsNullOrWhiteSpace(attachmentFile));

            DateTime? weekToStartEmailing = this.DateToSendIndividualEmails.SelectedDate;
            allFieldsValid = (allFieldsValid && weekToStartEmailing.HasValue);
            DateTime? weekToBulkEmail = this.DateToSendMassEmail.SelectedDate;
            if(allFieldsValid)
            {
                allFieldsValid = (allFieldsValid && weekToBulkEmail.HasValue && weekToBulkEmail.Value > weekToStartEmailing.Value);
            }
            string teamsMessagesSaveFile = TeamsMessagesFileNameTextBlock.Text;
            allFieldsValid = (allFieldsValid && !String.IsNullOrWhiteSpace(teamsMessagesSaveFile));
            string emailSender = EmailSenderTextBox.Text;
            allFieldsValid = (allFieldsValid && !String.IsNullOrWhiteSpace(emailSender));

            IEmailSender senderToUse = new EmailSenderOutlook(attachmentFile);
            if(allFieldsValid)
            {
                try
                {
                    LateTimesheetContactor contactor = new LateTimesheetContactor(timesheetRepository,
                                                            senderToUse,
                                                            weekToStartEmailing.Value,
                                                            weekToBulkEmail.Value,
                                                            timeOfReportCreation,
                                                            teamsMessagesSaveFile,
                                                            emailSender);                
                    contactor.ContactLateTimesheetEmployees();
                    string message = $"The emails have been sent and the teams messages to send have been saved to {teamsMessagesSaveFile}";
                    string title = "Success";
                    MessageBox.Show(message, title);
                }
                catch
                {
                    string message = "Something went wrong. Please check your Outlook sent emails to see if any emails were sent before retrying to avoid repeat emails.";
                    string title = "Error";
                    MessageBox.Show(message, title);
                }
                App.ParentWindowRef.Close();
            }
            else
            {
                    string message = "Please ensure all fields are filled out correctly.";
                    string title = "More Information Needed";
                    MessageBox.Show(message, title);
            }
        }
    }
}