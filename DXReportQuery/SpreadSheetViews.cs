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
            string sheetTitle = string.Format("{0}至{1}问题登记回访报表(一线)", Config.beginTime, Config.endTime);
            int sheetRowCounts = 0;             //表单内容当前行数

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
            DjwtSheetStyleInit(workbook);

            workbook.DocumentSettings.R1C1ReferenceStyle = true;

            Worksheet worksheet = SpreadView.GetWorkSheet(workbook, sheetName);
            worksheet.ActiveView.ShowGridlines = false;
            Range sheetTitleRange = worksheet.Range.FromLTRB(sheetRowCounts, sheetRowCounts, 8, sheetRowCounts);
            worksheet.MergeCells(sheetTitleRange);
            sheetTitleRange.Style = workbook.Styles["myDjwtSheetTitleStyle"];
            sheetTitleRange.SetValueFromText(sheetTitle);
            sheetRowCounts += 1;

            List<string> sheetTableHeadList = new List<string> { "行业", "工号", "姓名", "问题登记量", "回访数", "需求登记量", "回访率", "关闭（解决）率", "平均回访周期" };
            for(int i = 0; i < sheetTableHeadList.Count; i++)
            {
                Range sheetTableHeadRange = worksheet.Range.FromLTRB(i, sheetRowCounts, i, sheetRowCounts);
                sheetTableHeadRange.SetValueFromText(sheetTableHeadList[i].ToString());
                sheetTableHeadRange.Style = workbook.Styles["myDjwtSheetHeadSytle"];
                if (i >= 6)
                {
                    sheetTableHeadRange.Style = workbook.Styles["Output"];
                }
            }
            sheetRowCounts += 1;

            foreach (KeyValuePair<string, int> kv in branchCount)
            {
                DataRow[] dataRows = djwtDataTable.Select($@"ver='{kv.Key}'");
                foreach(DataRow dr in dataRows)
                {
                    for (int i = 0; i < djwtDataTable.Columns.Count; i++)
                    {
                        Range sheetNormal = worksheet.Range.FromLTRB(i, sheetRowCounts, i, sheetRowCounts);
                        sheetNormal.SetValueFromText(dr[i].ToString());
                        sheetNormal.Style = workbook.Styles["myDjwtSheetNormalSytle"];

                        if (i == 6)
                        {
                            sheetNormal.NumberFormat = "0.00%";
                        }
                    }
                    sheetRowCounts += 1;
                }

                for(int i = 0; i < djwtDataTable.Columns.Count; i++)
                {
                    if(i == 0)
                    {
                        worksheet.MergeCells(worksheet.Range.FromLTRB(0, sheetRowCounts - kv.Value, 0, sheetRowCounts));
                    }

                    if (i == 1)
                    {
                        Range sheetSubTotal = worksheet.Range.FromLTRB(1, sheetRowCounts, 1, sheetRowCounts);
                        sheetSubTotal.SetValueFromText("小计：");
                        sheetSubTotal.Style = workbook.Styles["myDjwtSheetSubTotalSytle"];
                    }

                    if ( i>=3 & i<=6)
                    {
                        Range sheetSubTotal = worksheet.Range.FromLTRB(i, sheetRowCounts, i, sheetRowCounts);
                        sheetSubTotal.Formula = $"=SUM(R[-{kv.Value}]C:R[-1]C)";
                        sheetSubTotal.Style = workbook.Styles["myDjwtSheetSubTotalSytle"];

                        if (i == 6)
                        {
                            sheetSubTotal.Formula = $"=AVERAGE(R[-{kv.Value}]C:R[-1]C)";
                            sheetSubTotal.NumberFormat = "0.00%";
                        }
                        
                    }

                }
                sheetRowCounts += 1;
                
            }

            workbook.DocumentSettings.R1C1ReferenceStyle = false;
            frmMainView.frmMainForm.ssQueryResultView.EndUpdate();
        }

        public static void WtgblView()
        {
            frmMainView.frmMainForm.ssQueryResultView.BeginUpdate();

            string sheetName = "";
            string sheetTitle = string.Format("{0}至{1}", Config.beginTime, Config.endTime);
            int sheetRowCounts = 0;

            frmMainView.frmMainForm.ssQueryResultView.BeginUpdate();

        }
    }

}
