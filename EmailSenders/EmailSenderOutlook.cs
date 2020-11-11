using Outlook = Microsoft.Office.Interop.Outlook;
using System;
using System.Collections.Generic;
using WPFLateTimesheetsApp.Models;

namespace WPFLateTimesheetsApp.EmailSenders
{
    public class EmailSenderOutlook : IEmailSender
    {
        private Outlook.Application app;

        private string _attachmentPath;
        public EmailSenderOutlook(string attachmentFilePath)
        {
            app = new Outlook.Application();
            _attachmentPath = attachmentFilePath;
        }

        public void SendLateTimesheetEmail(string lateEmployeeEmail, string peopleManagerEmail, string emailBody)
        {
            Outlook.MailItem timesheetEmail = (Outlook.MailItem)app.CreateItem(Outlook.OlItemType.olMailItem);
            timesheetEmail.To = lateEmployeeEmail;
            Console.WriteLine(lateEmployeeEmail + " " + peopleManagerEmail);
            if(!string.IsNullOrWhiteSpace(peopleManagerEmail) && !Int32.TryParse(peopleManagerEmail, out int result))
            {
                timesheetEmail.BCC = peopleManagerEmail;
            }
            timesheetEmail.Subject = "ACTION NEEDED: GTE Timesheet Error";
            timesheetEmail.HTMLBody = emailBody;
            timesheetEmail.Attachments.Add(_attachmentPath, Outlook.OlAttachmentType.olByValue, 1, "GTE Helpful Hint.pdf");
            timesheetEmail.Send();
        }

        public void SendLateTimesheetMassEmail(IEnumerable<LateTimesheet> bulkContactTimesheets, string emailBody)
        {
            Outlook.MailItem timesheetEmail = (Outlook.MailItem)app.CreateItem(Outlook.OlItemType.olMailItem);
            bool firstEmail = true;
            foreach(var lateSheet in bulkContactTimesheets)
            {
                // timesheetEmail.Recipients.Add(lateSheet.EmployeeEmail);
                if(firstEmail)
                {
                    timesheetEmail.BCC = lateSheet.EmployeeEmail;
                    firstEmail = false;
                }
                else
                {
                    timesheetEmail.BCC += ";" + lateSheet.EmployeeEmail;
                }
            }
            timesheetEmail.Subject = "ACTION NEEDED: GTE Timesheet Error";
            timesheetEmail.HTMLBody = emailBody;
            timesheetEmail.Attachments.Add(_attachmentPath, Outlook.OlAttachmentType.olByValue, 1, "GTE Helpful Hint.pdf");
            timesheetEmail.Send();
        }
    }
}