using System.Collections.Generic;
using WPFLateTimesheetsApp.Models;

namespace WPFLateTimesheetsApp.EmailSenders
{
    public interface IEmailSender
    {
        void SendLateTimesheetEmail(string lateEmployeeEmail, string peopleManagerEmail,
                                            string emailBody);
        void SendLateTimesheetMassEmail(IEnumerable<LateTimesheet> bulkContactTimesheets, string emailBody);
    }
}


