using System;
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
using WPFLateTimesheetsApp.LateTimesheetsRepositories;
using WPFLateTimesheetsApp.Controllers;
using WPFLateTimesheetsApp.EmailSenders;

namespace WPFLateTimesheetsApp
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            App.ParentWindowRef = this;
            this.WindowFrame.Navigate(new InputFormPage());
        }

        // public void DisplayTimesheetSelectionWindow(object sender, RoutedEventArgs e)
        // {
        //     OpenFileDialog fileSelectionDialog = new Microsoft.Win32.OpenFileDialog();

        //     Nullable<bool> fileSelected = fileSelectionDialog.ShowDialog();

        //     if(fileSelected.HasValue && fileSelected.Value)
        //     { 
        //         this.SelectedFileTextBlock.Text = fileSelectionDialog.FileName;
        //     }
        // }

        // public void SendEmailsAndDisplayUsersToMessage(object sender, RoutedEventArgs e)
        // {
        //     string timeOfReportCreation = this.ReportCreationTimeTextBox.Text; 
        //     string timeSheetFile = this.SelectedFileTextBlock.Text; 
        //     ILateTimesheetRepository timesheetRepository = new LateTimesheetsCsvRepository(timeSheetFile);
        //     DateTime? weekToStartEmailing = this.DateToSendIndividualEmails.SelectedDate;
        //     DateTime? weekToBulkEmail = this.DateToSendMassEmail.SelectedDate;

        //     if(weekToStartEmailing.HasValue && weekToBulkEmail.HasValue)
        //     {
        //         LateTimesheetContactor contactor = new LateTimesheetContactor(timesheetRepository, new EmailSenderOutlook(), weekToStartEmailing.Value, weekToBulkEmail.Value, timeOfReportCreation);
        //         contactor.ContactLateTimesheetEmployees();

        //     }
        // }
    }
}
