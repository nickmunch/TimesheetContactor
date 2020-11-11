using System.Collections.Generic;
using WPFLateTimesheetsApp.Models;

namespace WPFLateTimesheetsApp.LateTimesheetsRepositories
{
    public interface ILateTimesheetRepository
    {
        IEnumerable<LateTimesheet> GetLateTimesheets();
    }
}
