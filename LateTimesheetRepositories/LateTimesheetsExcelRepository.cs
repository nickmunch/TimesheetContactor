using System;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using WPFLateTimesheetsApp.Models;

namespace WPFLateTimesheetsApp.LateTimesheetsRepositories
{
    public class LateTimesheetsExcelRepository : ILateTimesheetRepository
    {
        private string _lateTimeSheetFilePath;
        public LateTimesheetsExcelRepository(string filePath)
        {
            _lateTimeSheetFilePath = filePath;
        }

        public IEnumerable<LateTimesheet> GetLateTimesheets()
        {
            List<LateTimesheet> lateTimesheets = new List<LateTimesheet>();

            if(File.Exists(_lateTimeSheetFilePath))
            {
                Excel.Application excelApp = new Excel.Application();
                Excel.Workbook timesheetsWorkbook = excelApp.Workbooks.Open(_lateTimeSheetFilePath);
                Excel.Worksheet timesheetsWorksheet = (Excel.Worksheet) timesheetsWorkbook.Sheets[1];
                Excel.Range timesheetRange = timesheetsWorksheet.UsedRange;

                int numRows = timesheetRange.Rows.Count;

                for(int i = 2; i <= numRows; i++)
                {
                        var currentTimesheet = new LateTimesheet();
                        currentTimesheet.FullName = (string) (timesheetRange.Cells[i, 1] as Excel.Range).Text;
                        currentTimesheet.Rank = (string) (timesheetRange.Cells[i, 2] as Excel.Range).Text;
                        currentTimesheet.EmployeeId = (string) (timesheetRange.Cells[i, 3] as Excel.Range).Text;
                        currentTimesheet.CycleEndDate = Convert.ToDateTime((string) (timesheetRange.Cells[i, 4] as Excel.Range).Text);
                        currentTimesheet.Speciality = (string) (timesheetRange.Cells[i, 5] as Excel.Range).Text;
                        currentTimesheet.TimeSheetStatus = (string) (timesheetRange.Cells[i, 6] as Excel.Range).Text;
                        currentTimesheet.EmployeeEmail = (string) (timesheetRange.Cells[i, 7] as Excel.Range).Text;
                        currentTimesheet.PeopleManager = (string) (timesheetRange.Cells[i, 8] as Excel.Range).Text;
                        currentTimesheet.PeopleManagerEmail = (string) (timesheetRange.Cells[i, 9] as Excel.Range).Text;
                        currentTimesheet.CurrentMarketUnit = (string) (timesheetRange.Cells[i, 10] as Excel.Range).Text;
                        currentTimesheet.Account = (string) (timesheetRange.Cells[i, 11] as Excel.Range).Text;
                        lateTimesheets.Add(currentTimesheet);
                }

                Marshal.ReleaseComObject(timesheetRange);
                Marshal.ReleaseComObject(timesheetsWorksheet);

                timesheetsWorkbook.Close();
                Marshal.ReleaseComObject(timesheetsWorkbook);

                excelApp.Quit();
                Marshal.ReleaseComObject(excelApp);
            }

            return(lateTimesheets);
        }
    }
}
