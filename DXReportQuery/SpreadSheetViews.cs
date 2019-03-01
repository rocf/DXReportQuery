using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data;
using DevExpress.Spreadsheet;
using static DXReportQuery.SpreadSheetStyles;

namespace DXReportQuery
{
    partial class SpreadView
    {
        private static Worksheet GetWorkSheet(IWorkbook workbook, string sheetName)
        {

            if (workbook.Worksheets.Contains("Sheet1"))
            {
                Worksheet worksheet = workbook.Worksheets["Sheet1"];
                worksheet.Name = sheetName;
                return worksheet;
            }

            if (!workbook.Worksheets.Contains(sheetName))
            {
                Worksheet worksheet = workbook.Worksheets.Add(sheetName);
                return worksheet;
            }

            Worksheet newworksheet = workbook.Worksheets[sheetName];
            return newworksheet;         
        }
        public static void DjwtView()
        {
            frmMainView.frmMainForm.ssQueryResultView.BeginUpdate();

            string sheetName = "问题登记";
            string sheetHead = string.Format("{0}至{1}问题登记回访报表(一线)", Config.beginTime, Config.endTime);

            DataTable djwtDataTable = QueryResults.DjwtQuery();
            Dictionary<string, int> branchCount = new Dictionary<string, int>();

            var queryCountResult = from djwt in djwtDataTable.AsEnumerable()
                                   group djwt by new { ver = djwt.Field<string>("ver") }
                                   into g
                                   select new
                                   {
                                       g.Key.ver,
                                       count = g.Count()
                                   };

            if (queryCountResult.ToList().Count > 0)
            {
                queryCountResult.ToList().ForEach(q => branchCount.Add(q.ver, q.count));
            }
            
            IWorkbook workbook = frmMainView.frmMainForm.ssQueryResultView.Document;
            Worksheet worksheet = SpreadView.GetWorkSheet(workbook, sheetName);
            Range sheetHeadRange = worksheet.Range["A1:H1"];
            worksheet.MergeCells(sheetHeadRange);
            sheetHeadRange.Style = DjwtSheetStyle.DjwtSheetTitleStyle(workbook);
            sheetHeadRange.SetValueFromText(sheetHead);


            frmMainView.frmMainForm.ssQueryResultView.EndUpdate();
        }
    }
}
