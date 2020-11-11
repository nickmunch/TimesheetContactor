using System;

namespace WPFLateTimesheetsApp.Models
{
    public class LateTimesheet
    {
        public string FullName { get; set; }
        public string Rank { get; set; }
        public string EmployeeId {get; set;}
        public DateTime CycleEndDate { get; set; }

        public string Speciality {get; set;}

        public string TimeSheetStatus {get; set;}

        public string EmployeeEmail {get; set;}

        public string PeopleManager {get; set;}

        public string PeopleManagerEmail {get; set;}

        public string CurrentMarketUnit {get; set;}

        public string Account {get; set;}

        public string Comments {get; set;}

    }
}