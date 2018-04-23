using Microsoft.Office.Interop.Excel;
using System.Collections.Generic;

namespace CefSharp.MinimalExample.WinForms
{
    public class Excel
    {
        public static List<ExcelRecord> ReadFile(string FilePath)
        {
            var result = new List<ExcelRecord>();

            var file = new Application();

            var wb = file.Application.Workbooks.Open(FilePath);
            //wb.Activate();
            var ws = wb.Worksheets[1] as Worksheet;
            
            for (long i = 2; i < ws.UsedRange.Rows.Count; i++)
            {
                var row = ws.UsedRange.Rows.EntireRow;
                result.Add(row.map(i));
            }

            wb.Close(false);
            file.Quit();
            return result;
        }

        public static void UpdateRecord(string FilePath, ExcelRecord record)
        {
            var file = new Application();

            var wb = file.Application.Workbooks.Open(FilePath);
            //wb.Activate();
            var ws = wb.Worksheets[1] as Worksheet;

            ws.Cells[record.RowNumber, "J"].Value = record.Präsens;
            ws.Cells[record.RowNumber, "K"].Value = record.Präteritum;
            ws.Cells[record.RowNumber, "E"].Value = record.Partizip;
            ws.Cells[record.RowNumber, "F"].Value = record.Perfekt;



            wb.Save();
            wb.Close();
            file.Quit();
        }
    }
}
