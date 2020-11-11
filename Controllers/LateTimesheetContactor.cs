using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using WPFLateTimesheetsApp.Models;
using System.Net;
using WPFLateTimesheetsApp.LateTimesheetsRepositories;
using WPFLateTimesheetsApp.EmailSenders;

namespace WPFLateTimesheetsApp.Controllers
{
    public class LateTimesheetContactor
    {
        private ILateTimesheetRepository _repository;

        private IEmailSender _emailSender;
        private DateTime _weekToEmail;

        private string _reportCreationTimeString;

        private DateTime _weekToBulkUpdate;

        private string _emailOpeningString;
        private string _emailClosingString;

        private string _teamsMessagesSaveFile;
        private string _senderName;
        public LateTimesheetContactor(ILateTimesheetRepository timesheetRepo, IEmailSender emailSender, DateTime emailWeek, DateTime bulkUpdateWeek,
                                      string reportCreationTime, string locationToSaveTeamsMessages, string senderName)
        {
            _repository = timesheetRepo;
            _emailSender = emailSender;
            _weekToEmail = emailWeek;
            _weekToBulkUpdate = bulkUpdateWeek;
            _reportCreationTimeString = reportCreationTime;
            _teamsMessagesSaveFile = locationToSaveTeamsMessages;
            _senderName = senderName;
            _emailClosingString = null;
            _emailOpeningString = null;
        }

        public void ContactLateTimesheetEmployees()
        {
            ICollection<LateTimesheet> bulkContactTimesheets = new List<LateTimesheet>();
            ICollection<List<LateTimesheet>> teamsMessageTimesheets = new List<List<LateTimesheet>>();
            IEnumerable<LateTimesheet> lateTimesheets = _repository.GetLateTimesheets();
            var timeSheetGroups = lateTimesheets.OrderBy(t => t.CycleEndDate).GroupBy(t => t.EmployeeId);

            

            foreach(var employeeLateTimeSheetGroup in timeSheetGroups)
            {
                var earliestLateTimesheet = employeeLateTimeSheetGroup.First();

                if(earliestLateTimesheet.CycleEndDate < _weekToEmail)
                {
                    // SendLateTimesheetMessage(employeeLateTimeSheetGroup);
                    teamsMessageTimesheets.Add(employeeLateTimeSheetGroup.ToList());
                }
                else if(earliestLateTimesheet.CycleEndDate < _weekToBulkUpdate)
                {
                    SendLateTimesheetEmail(employeeLateTimeSheetGroup);
                }
                else if(earliestLateTimesheet.CycleEndDate == _weekToBulkUpdate)
                {
                    bulkContactTimesheets.Add(earliestLateTimesheet);
                }
            }
            if(teamsMessageTimesheets.Count > 0)
            {
                CreateAndSaveTeamsMessages(teamsMessageTimesheets);
            }
            if(bulkContactTimesheets.Count > 0)
            {
                SendMassEmail(bulkContactTimesheets);
            }
            
        }

        private void CreateAndSaveTeamsMessages(IEnumerable<List<LateTimesheet>> teamsMessagesTimesheets)
        {
            StringBuilder teamsMessages = new StringBuilder();
            foreach(var timesheetsGroup in teamsMessagesTimesheets)
            {
                LateTimesheet firstLateTimesheet = timesheetsGroup.First();
                teamsMessages.Append($"Messaging {firstLateTimesheet.FullName}\n");
                teamsMessages.Append($"Hello,\n\nAccording to a report generated with data pulled from {_reportCreationTimeString},");
                teamsMessages.Append(" you are missing your timesheet(s) for the following week end dates. Please submit your timesheet ASAP and let me know when this is complete.\n\n");
                foreach(var sheet in timesheetsGroup)
                {
                    teamsMessages.Append($"Name: {sheet.FullName}\n");
                    teamsMessages.Append($"Missing Timesheet: {sheet.CycleEndDate}\n");
                    teamsMessages.Append($"Status of Timesheet: {sheet.TimeSheetStatus}\n\n");
                }
                teamsMessages.Append("Please make sure to \"Submit\" the timesheet and not just \"Save\".\n\n");
                teamsMessages.Append("Thank you\n");
                teamsMessages.Append($"{_senderName}\n");
                teamsMessages.Append("------------------------------------------------------------\n");
            }

            System.IO.File.WriteAllText(_teamsMessagesSaveFile, teamsMessages.ToString());
        }

        private void SendMassEmail(IEnumerable<LateTimesheet> massEmailTimesheets)
        {
            _emailSender.SendLateTimesheetMassEmail(massEmailTimesheets, CreateEmailBody());
        }

        private void SendLateTimesheetMessage(IEnumerable<LateTimesheet> employeeLateTimesheets)
        {
            LateTimesheet employeeLatestTimesheet = employeeLateTimesheets.First();
            Console.WriteLine($" Messaging {employeeLatestTimesheet.FullName} for {employeeLatestTimesheet.TimeSheetStatus} timesheet for {employeeLatestTimesheet.CycleEndDate}");
            // Console.WriteLine("-----------------------Table--------------------");
            // var messageBuilder = new StringBuilder();
            // CreateEmployeeMissingTimesheetTable(messageBuilder, employeeLateTimesheets);
            // Console.WriteLine(messageBuilder.ToString());
            // Console.WriteLine("------------------------------------------------");
        }
        
        private void SendLateTimesheetEmail(IEnumerable<LateTimesheet> employeeLateTimesheets)
        {
            LateTimesheet employeeLatestTimesheet = employeeLateTimesheets.First();
            var messageBuilder = new StringBuilder();
            Console.WriteLine($" Emailing {employeeLatestTimesheet.FullName} for {employeeLatestTimesheet.TimeSheetStatus} timesheet for {employeeLatestTimesheet.CycleEndDate}");
            // Console.WriteLine("------------------------------------------------");
            _emailSender.SendLateTimesheetEmail(employeeLatestTimesheet.EmployeeEmail, employeeLatestTimesheet.PeopleManagerEmail, CreateEmailBodyWithTable(employeeLateTimesheets));
        }

        private string CreateEmailBodyWithTable(IEnumerable<LateTimesheet> employeeLateTimesheets)
        {
            var messageBuilder = new StringBuilder();
            messageBuilder.Append(CreateEmailOpening());
            CreateEmployeeMissingTimesheetTable(messageBuilder, employeeLateTimesheets);
            messageBuilder.Append(CreateEmailClosing());
            return(messageBuilder.ToString());
        }

        private string CreateEmailBody()
        {
            var messageBuilder = new StringBuilder();
            messageBuilder.Append(CreateEmailOpening());
            messageBuilder.Append(CreateEmailClosing());
            return(messageBuilder.ToString());
        }

        private string CreateEmailOpening()
        {
            if(_emailOpeningString == null)
            {
                var emailOpeningStringBuilder = new StringBuilder();
                if(File.Exists("./emailOpening.html"))
                {
                    using(var emailHeaderReader = new StreamReader("./emailOpening.html"))
                    {
                        emailOpeningStringBuilder.Append(emailHeaderReader.ReadToEnd());
                    }
                }
                emailOpeningStringBuilder.Append($@"<div class=WordSection1>
                    <p class=MsoNormal><i>Hello,&nbsp;&nbsp; <o:p></o:p></i></p>
                    <p class=MsoNormal><i>&nbsp;&nbsp;&nbsp;<o:p></o:p></i></p>
                    <p class=MsoNormal><i>According to a report generated with data pulled from <b>");
                emailOpeningStringBuilder.Append(_reportCreationTimeString); //(20-Oct, 2020, 6:00 PM EST)
                emailOpeningStringBuilder.Append(@"</b>, you are missing your timesheet(s) for the following week end dates. Please submit your timesheet ASAP
                        and let me know when this is complete.&nbsp;&nbsp; <o:p></o:p></i></p>");
                _emailOpeningString = emailOpeningStringBuilder.ToString();
            }
            return(_emailOpeningString);
        }

        private string CreateEmailClosing()
        {
            if(_emailClosingString == null)
            {
                if(File.Exists("./emailClosing.html"))
                {
                    using(var emailClosingReader = new StreamReader("./emailClosing.html"))
                    {
                        _emailClosingString = emailClosingReader.ReadToEnd();
                    }
                }
                _emailClosingString += _senderName + @"</p></div></body></html>";
            }

            return(_emailClosingString);
        }

        // private string ReadSignature() 
        // { 
        //     string appDataDir = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + "\\Microsoft\\Signatures"; 
        //     string signature = string.Empty;
        //     DirectoryInfo diInfo = new DirectoryInfo(appDataDir); 
        //     if(diInfo.Exists) 
        //     { 
        //         FileInfo[] fiSignature = diInfo.GetFiles("*.htm"); 
        //         if (fiSignature.Length > 0) 
        //         { 
        //             StreamReader sr = new StreamReader(fiSignature[0].FullName, Encoding.Default); 
                    
        //             signature = sr.ReadToEnd(); 
        //             if (!string.IsNullOrEmpty(signature)) 
        //             {
        //                 string fileName = fiSignature[0].Name.Replace(fiSignature[0].Extension, string.Empty);
        //                 Console.WriteLine(fileName);
        //                 signature = signature.Replace(Uri.EscapeUriString(fileName) + "_files/", appDataDir + "\\" + fileName + "_files\\"); 
        //             }
        //         } 

        //     }
        //     return signature; 
        // } 

        private void CreateEmployeeMissingTimesheetTable(StringBuilder employeeHtmlTable, IEnumerable<LateTimesheet> employeeLateTimesheets)
        {
            employeeHtmlTable.Append("<table class=MsoNormalTable border=0 cellspacing=0 cellpadding=0 width=379 style='width:284.6pt;border-collapse:collapse'>");
            CreateNewLateTimesheetTableRow(employeeHtmlTable, "Full Name", "Cycle End Date", "Time Sheet Status", true);
            foreach(var timesheet in employeeLateTimesheets)
            {
                CreateNewLateTimesheetTableRow(employeeHtmlTable, timesheet.FullName, timesheet.CycleEndDate.ToString("MM/dd/yyyy"), timesheet.TimeSheetStatus, false);
            }

            employeeHtmlTable.Append("</table>");
        }

        private void CreateNewLateTimesheetTableRow(StringBuilder employeeHtmlTable, string fullName, string cycleEndDate, string timesheetStatus, bool isHeader)
        {
            employeeHtmlTable.Append("<tr style='height:21.5pt'>");
            AppendTableElement(employeeHtmlTable, fullName, isHeader);
            AppendTableElement(employeeHtmlTable, cycleEndDate, isHeader);
            AppendTableElement(employeeHtmlTable, timesheetStatus, isHeader);
            employeeHtmlTable.Append("</tr>");
        }

        private void AppendTableElement(StringBuilder employeeTable, string elementToAdd, bool isHeader)
        {
            string tableCellBeginning;
            if(isHeader)
            {
                tableCellBeginning = @"<td width=156 nowrap valign=bottom
                    style='width:117.0pt;border-top:solid #9BC2E6 1.0pt;border-left:none;border-bottom:solid #9BC2E6 1.0pt;border-right:none;background:#5B9BD5;padding:0in 5.4pt 0in 5.4pt;height:30.5pt'>
                    <p class=MsoNormal><b><span style='color:white'>";
            }
            else
            {
                tableCellBeginning = @"<td width=156 nowrap valign=bottom
                    style='width:117.0pt;background:#DDEBF7;padding:0in 5.4pt 0in 5.4pt;height:21.5pt'>
                    <p class=MsoNormal><span style='color:black'>";
            }
            string tableCellEnd = @"<o:p></o:p></span></b></p></td>";

            employeeTable.Append(tableCellBeginning);
            employeeTable.Append(elementToAdd);
            employeeTable.Append(tableCellEnd);
        }
    }
}