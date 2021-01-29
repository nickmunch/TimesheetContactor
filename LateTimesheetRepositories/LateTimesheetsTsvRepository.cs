using System;
using System.IO;
using System.Collections.Generic;
using WPFLateTimesheetsApp.Models;

namespace WPFLateTimesheetsApp.LateTimesheetsRepositories
{
    public class LateTimesheetsTsvRepository : ILateTimesheetRepository
    {
        private string _lateTimeSheetFilePath;
        public LateTimesheetsTsvRepository(string filePath)
        {
            _lateTimeSheetFilePath = filePath;
        }

        public IEnumerable<LateTimesheet> GetLateTimesheets()
        {
            List<LateTimesheet> lateTimesheets = new List<LateTimesheet>();

            if(File.Exists(_lateTimeSheetFilePath))
            {
                using(var timesheetReader = new StreamReader(_lateTimeSheetFilePath))
                {
                    string headerLine = timesheetReader.ReadLine();
                    string line;
                    while((line = timesheetReader.ReadLine()) != null)
                    {
                        var lineColumns = line.Split('\t');
                        var currentTimesheet = new LateTimesheet();
                        currentTimesheet.FullName = lineColumns[0];
                        currentTimesheet.Rank = lineColumns[1];
                        currentTimesheet.EmployeeId = lineColumns[2];
                        currentTimesheet.CycleEndDate = Convert.ToDateTime(lineColumns[3]);
                        // Column 5 is the Prod Unit
                        currentTimesheet.Speciality = lineColumns[5];
                        currentTimesheet.TimeSheetStatus = lineColumns[6];
                        currentTimesheet.EmployeeEmail = lineColumns[7];
                        currentTimesheet.PeopleManager = lineColumns[8];
                        currentTimesheet.PeopleManagerEmail = lineColumns[9];
                        lateTimesheets.Add(currentTimesheet);
                    }
                }
            }

            return(lateTimesheets);
        }
    }
}
