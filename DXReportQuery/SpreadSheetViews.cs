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
            workbook.Worksheets.ActiveWorksheet = workbook.Worksheets[sheetName];
            worksheet.ActiveView.ShowGridlines = false;
            Range sheetTitleRange = worksheet.Range.FromLTRB(0, sheetRowCounts, 8, sheetRowCounts);
            worksheet.MergeCells(sheetTitleRange);
           // sheetTitleRange.Style = workbook.Styles["myDjwtSheetTitleStyle"];
            sheetTitleRange.SetValueFromText(sheetTitle);
            sheetRowCounts += 1;

            List<string> sheetTableHeadList = new List<string> { "行业", "工号", "姓名", "问题登记量", "回访数", "需求登记量", "回访率", "关闭（解决）率", "平均回访周期" };
            for(int i = 0; i < sheetTableHeadList.Count; i++)
            {
                Range sheetTableHeadRange = worksheet.Range.FromLTRB(i, sheetRowCounts, i, sheetRowCounts);
                sheetTableHeadRange.SetValueFromText(sheetTableHeadList[i].ToString());
                //sheetTableHeadRange.Style = workbook.Styles["myDjwtSheetHeadSytle"];
                if (i >= 6)
                {
                   // sheetTableHeadRange.Style = workbook.Styles["Output"];
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
                        //sheetNormal.Style = workbook.Styles["myDjwtSheetNormalSytle"];

                        if (i == 6)
                        {
                            sheetNormal.NumberFormat = "0.00%";
                        }
                    }
                    sheetRowCounts += 1;
                }

                for(int i = 0; i < djwtDataTable.Columns.Count; i++)
                {
                    Range sheetSubTotal = worksheet.Range.FromLTRB(i, sheetRowCounts, i, sheetRowCounts);

                    if (i == 0)
                    {
                        sheetSubTotal = worksheet.Range.FromLTRB(0, sheetRowCounts - kv.Value, 0, sheetRowCounts);
                        worksheet.MergeCells(sheetSubTotal);
                    }

                    if (i == 1)
                    {
                        
                        sheetSubTotal.SetValueFromText("小计：");
                        //sheetSubTotal.Style = workbook.Styles["myDjwtSheetSubTotalSytle"];
                    }

                    if ( i>=3 & i<=6)
                    {
                        sheetSubTotal.Formula = $"=SUM(R[-{kv.Value}]C:R[-1]C)";
                       // sheetSubTotal.Style = workbook.Styles["myDjwtSheetSubTotalSytle"];

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

            string sheetName = "问题关闭率";
            string sheetTitle1 = string.Format("{0}至{1}", Config.beginTime, Config.endTime);
            string sheetTitle2 = "整体关闭率统计";
            string sheetTitle3 = "vip问题关闭率统计";
            string sheetTitle4 = "处理中问题处理完成情况统计";
            string sheetTitle5 = "问题效能分析报表（全部）";

            int sheetRowCounts = 0;

            DataTable ztgblDataTable= QueryResults.ZtgblQuery();
            int ztgblRowCount = ztgblDataTable.Rows.Count;

            DataTable vipWtgblDataTable = QueryResults.VIPWtgblQuery();
            int vipWtgblRowCount = vipWtgblDataTable.Rows.Count;

            DataTable clzWtclDataTable = QueryResults.ClzWtclQuery();
            int clzWtclRowCount = clzWtclDataTable.Rows.Count;


            IWorkbook workbook = frmMainView.frmMainForm.ssQueryResultView.Document;
            workbook.DocumentSettings.R1C1ReferenceStyle = true;

            Worksheet worksheet = SpreadView.GetWorkSheet(workbook, sheetName);
            workbook.Worksheets.ActiveWorksheet = workbook.Worksheets[sheetName];
            worksheet.ActiveView.ShowGridlines = false;

            Range sheetTitle1Range = worksheet.Range.FromLTRB(0, sheetRowCounts, 14, sheetRowCounts);
            worksheet.MergeCells(sheetTitle1Range);
            sheetTitle1Range.SetValueFromText(sheetTitle1);
            sheetRowCounts += 1;

            Range sheetTitle2Range = worksheet.Range.FromLTRB(0, sheetRowCounts, 14, sheetRowCounts);
            worksheet.MergeCells(sheetTitle2Range);
            sheetTitle2Range.SetValueFromText(sheetTitle2);
            sheetRowCounts += 1;

            List<string> sheetTable1HeadList = new List<string> { "部门", "问题总数", "上周问题总数", "环比", "环比增长率", "全部问题(过滤付费)", "无状态问题(过滤付费)", "付费问题量", "待用户确认", "处理中", "待处理", "关闭", "关闭率", "上周关闭率", "同比" };

            for (int i = 0; i < sheetTable1HeadList.Count; i++)
            {
                Range sheetTable1HeadRange = worksheet.Range.FromLTRB(i, sheetRowCounts, i, sheetRowCounts);
                sheetTable1HeadRange.SetValueFromText(sheetTable1HeadList[i].ToString());
               // sheetTableHeadRange.Style = workbook.Styles[""];
            }
            sheetRowCounts += 1;

            foreach (DataRow dr in ztgblDataTable.AsEnumerable())
            {
                for (int i = 0; i < ztgblDataTable.Columns.Count; i++)
                {
                    Range sheetNormal = worksheet.Range.FromLTRB(i, sheetRowCounts, i, sheetRowCounts);
                    sheetNormal.SetValueFromText(dr[i].ToString());
                    //sheetNormal.Style = workbook.Styles[""];
                    if(i == 4 || (i >= 12 & i <= 14))
                    {
                        sheetNormal.NumberFormat = "0.00%";

                        if (i == 4)
                        {
                            sheetNormal.Formula = "RC[-1]/RC[-2]";
                        }

                        if (i == 14)
                        {
                            sheetNormal.Formula = "RC[-2] - RC[-1]";
                        }
                    }
                }
                sheetRowCounts += 1;
            }

            for (int i = 0; i < ztgblDataTable.Columns.Count; i++)            
            {
                Range sheetTotal = worksheet.Range.FromLTRB(i, sheetRowCounts, i, sheetRowCounts);
                sheetTotal.Formula = $"SUM(R[-{ztgblRowCount}]C:R[-1]C)";

                if (i == 0)
                {
                    sheetTotal.SetValueFromText("合计：");
                }

                if (i == 4)
                {
                    sheetTotal.Formula = "RC[-1]/RC[-2]";
                }

                if (i >= 12 & i <= 14)
                {
                    sheetTotal.Formula = $"AVERAGE(R[-{ztgblRowCount}]C:R[-1]C)";
                    sheetTotal.NumberFormat = "0.00%";
                }
            }
            sheetRowCounts += 2;

            Range sheetTitle3Range = worksheet.Range.FromLTRB(0, sheetRowCounts, 12, sheetRowCounts);
            worksheet.MergeCells(sheetTitle3Range);
            sheetTitle3Range.SetValueFromText(sheetTitle3);
            sheetRowCounts += 1;

            List<string> sheetTable2HeadList = new List<string> { "部门", "全部问题", "全部问题(过滤付费)", "无状态问题(过滤付费)", "付费问题量", "待用户确认", "处理中", "待处理", "关闭", "关闭率", "上周关闭率", "同比", "问题占比" };
            for (int i = 0; i < sheetTable2HeadList.Count; i++)
            {
                Range sheetTable2HeadRange = worksheet.Range.FromLTRB(i, sheetRowCounts, i, sheetRowCounts);
                sheetTable2HeadRange.SetValueFromText(sheetTable2HeadList[i].ToString());
                // sheetTableHeadRange.Style = workbook.Styles[""];
            }
            sheetRowCounts += 1;


            foreach (DataRow dr in vipWtgblDataTable.AsEnumerable())
            {
                for (int i = 0; i < vipWtgblDataTable.Columns.Count; i++)
                {
                    Range sheetNormal = worksheet.Range.FromLTRB(i, sheetRowCounts, i, sheetRowCounts);
                    sheetNormal.SetValueFromText(dr[i].ToString());
                    //sheetNormal.Style = workbook.Styles[""];

                    if (i >= 9 & i <= 12)
                    {                       
                        sheetNormal.NumberFormat = "0.00%";

                        if (i == 9)
                        {
                            sheetNormal.Formula = "RC[-1]/RC[-6]";
                        }

                        if (i == 11)
                        {
                            sheetNormal.Formula = "RC[-2] - RC[-1]";
                        }
                    }

                    
                }
                sheetRowCounts += 1;
            }

             for (int i = 0; i < vipWtgblDataTable.Columns.Count; i++)            
             {
                Range sheetTotal = worksheet.Range.FromLTRB(i, sheetRowCounts, i, sheetRowCounts);
                sheetTotal.Formula = $"SUM(R[-{vipWtgblRowCount}]C:R[-1]C)";

                if (i == 0)
                {
                    sheetTotal.SetValueFromText("合计：");
                }

                if (i >= 9 & i <= 12)
                {
                    sheetTotal.Formula = $"AVERAGE(R[-{ztgblRowCount}]C:R[-1]C)";
                    sheetTotal.NumberFormat = "0.00%";
                }
             }
            sheetRowCounts += 2;


            Range sheetTitle4Range = worksheet.Range.FromLTRB(0, sheetRowCounts, 11, sheetRowCounts);
            worksheet.MergeCells(sheetTitle4Range);
            sheetTitle4Range.SetValueFromText(sheetTitle4);
            sheetRowCounts += 1;

            List<string> sheetTable3HeadList = new List<string> { "部门", "问题总数", "全部问题(过滤付费)", "无状态问题(过滤付费)", "付费问题量", "待用户确认", "处理中", "待处理", "关闭", "关闭率", "上周关闭率", "同比" };
            for (int i = 0; i < sheetTable3HeadList.Count; i++)
            {
                Range sheetTable3HeadRange = worksheet.Range.FromLTRB(i, sheetRowCounts, i, sheetRowCounts);
                sheetTable3HeadRange.SetValueFromText(sheetTable3HeadList[i].ToString());

            }
            sheetRowCounts += 1;


            foreach(DataRow dr in clzWtclDataTable.AsEnumerable())
            {
                for(int i = 0; i < clzWtclDataTable.Columns.Count; i++)
                {
                    Range sheetNormal = worksheet.Range.FromLTRB(i, sheetRowCounts, i, sheetRowCounts);
                    sheetNormal.SetValueFromText(dr[i].ToString());

                    if (i >= 9 & i <= 11)
                    {
                        sheetNormal.NumberFormat = "0.00%";

                        if(i == 9)
                        {
                            sheetNormal.Formula = "RC[-1]/RC[-6]";
                        }

                        if (i == 11)
                        {
                            sheetNormal.Formula = "RC[-2] - RC[-1]";
                        }
                    }
                }
                sheetRowCounts += 1;
            }

            for (int i = 0; i < clzWtclDataTable.Columns.Count; i++)
            {
                Range sheetTotal = worksheet.Range.FromLTRB(i, sheetRowCounts, i, sheetRowCounts);
                sheetTotal.Formula = $"SUM(R[-{clzWtclRowCount}]C:R[-1]C)";

                if (i == 0)
                {
                    sheetTotal.SetValueFromText("合计：");
                }

                if (i >= 9 & i <= 11)
                {
                    sheetTotal.Formula = $"AVERAGE(R[-{clzWtclRowCount}]C:R[-1]C)";
                    sheetTotal.NumberFormat = "0.00%";

                    if (i == 11)
                    {
                        sheetTotal.Formula = "RC[-2] - RC[-1]";
                    }
                }
            }
            sheetRowCounts += 2;
        
            /*
            Range sheetTitle5Range = worksheet.Range.FromLTRB(0, sheetRowCounts, 5, sheetRowCounts);
            worksheet.MergeCells(sheetTitle5Range);
            sheetTitle5Range.SetValueFromText(sheetTitle5);
            sheetRowCounts += 1;

            List<string> sheetTable4HeadList = new List<string> {"行业", "平均响应速度", "平均解决周期", "平均关闭周期", "覆盖代理商数", "平均处理次数"};
            for(int i = 0; i < sheetTable4HeadList.Count; i++)
            {
                Range sheetTable4HeadRange = worksheet.Range.FromLTRB(i, sheetRowCounts, i, sheetRowCounts);
                sheetTable4HeadRange.SetValueFromText(sheetTable4HeadList[i].ToString());
            }
            sheetRowCounts += 1;

            */

            workbook.DocumentSettings.R1C1ReferenceStyle = false;
            frmMainView.frmMainForm.ssQueryResultView.EndUpdate();
        }

        public static void WtmjgblView()
        {


        }

        public static void QyxnView()
        {
            frmMainView.frmMainForm.ssQueryResultView.BeginUpdate();

            string sheetName = "区域效能";
            string sheetTitle = string.Format("{0}至{1}区域效能", Config.beginTime, Config.endTime);
            int sheetRowCounts = 0;             //表单内容当前行数

            IWorkbook workbook = frmMainView.frmMainForm.ssQueryResultView.Document;
            workbook.DocumentSettings.R1C1ReferenceStyle = true;

            Worksheet worksheet = SpreadView.GetWorkSheet(workbook, sheetName);
            workbook.Worksheets.ActiveWorksheet = workbook.Worksheets[sheetName];
            worksheet.ActiveView.ShowGridlines = false;
            Range sheetTitleRange = worksheet.Range.FromLTRB(0, sheetRowCounts, 12, sheetRowCounts);
            worksheet.MergeCells(sheetTitleRange);
            // sheetTitleRange.Style = workbook.Styles["myDjwtSheetTitleStyle"];
            sheetTitleRange.SetValueFromText(sheetTitle);
            sheetRowCounts += 1;

            List<string> sheetTableHeadList = new List<string> { "行业","负责人","省份","问题取样数","实际响应速度","实际解决周期","实际关闭周期","有效响应速度","有效解决周期","有效关闭周期","平均回复次数","响应超时","处理超时" };
            for (int i = 0; i < sheetTableHeadList.Count; i++)
            {
                Range sheetTableHeadRange = worksheet.Range.FromLTRB(i, sheetRowCounts, i, sheetRowCounts);
                sheetTableHeadRange.SetValueFromText(sheetTableHeadList[i].ToString());
                //sheetTableHeadRange.Style = workbook.Styles["myDjwtSheetHeadSytle"];

            }
            sheetRowCounts += 1;
            
            Dictionary<string, string> deptDic = new Dictionary<string, string>();
            deptDic.Add("1", "商超");
            deptDic.Add("C", "生鲜便利");
            deptDic.Add("8", "商锐");
            deptDic.Add("3", "专卖");
            deptDic.Add("H", "孕婴童");
            deptDic.Add("2", "餐饮");
            deptDic.Add("I", "星食客");
            deptDic.Add("6", "eshop");

            Dictionary<string, int> deptCountDic = new Dictionary<string, int>();
            string oldKey = "1";

            foreach (string key in deptDic.Keys)
            {
                DataTable QyxnDataTable = QueryResults.QyxnQuery(key);
                Dictionary<string, int> nameCount = new Dictionary<string, int>();
                List<int> sheetSubTotalRows = new List<int> { };
                string nowKey = key;

                deptCountDic.Add(key, QyxnDataTable.Rows.Count);

                var queryCountResult = from djwt in QyxnDataTable.AsEnumerable()
                                       group djwt by new { Name = djwt.Field<string>("Name") }
                                       into g
                                       select new
                                       {
                                           g.Key.Name,
                                           count = g.Count()
                                       };

                if (queryCountResult.ToList().Count > 0)
                {
                    queryCountResult.ToList().ForEach(q => nameCount.Add(q.Name, q.count));
                }

              

                foreach (KeyValuePair<string, int> kv in nameCount)
                {
                    DataRow[] dataRows = QyxnDataTable.Select($@"Name='{kv.Key}'");
                    
                    foreach (DataRow dr in dataRows)
                    {
                        for (int i = 0; i < QyxnDataTable.Columns.Count; i++)
                        {
                            Range sheetNormal = worksheet.Range.FromLTRB(i, sheetRowCounts, i, sheetRowCounts);
                            sheetNormal.SetValueFromText(dr[i].ToString());
                            //sheetNormal.Style = workbook.Styles["myDjwtSheetNormalSytle"];
                        }
                        sheetRowCounts += 1;
                    }

                    for (int i = 0; i < QyxnDataTable.Columns.Count; i++)
                    {
                        Range sheetSubTotal = worksheet.Range.FromLTRB(i, sheetRowCounts, i, sheetRowCounts);


                        if (i == 1)
                        {
                            sheetSubTotal = worksheet.Range.FromLTRB(i, sheetRowCounts - kv.Value, i, sheetRowCounts);
                            worksheet.MergeCells(sheetSubTotal);
                        }

                        if (i == 2)
                        {

                            sheetSubTotal.SetValueFromText("小计：");
                            //sheetSubTotal.Style = workbook.Styles["myDjwtSheetSubTotalSytle"];
                        }

                        if (i == 3)
                        {
                            sheetSubTotal.Formula = $"=SUM(R[-{kv.Value}]C:R[-1]C)";
                        }
                        if (i >= 4 & i <= 10)
                        {
                            sheetSubTotal.Formula = $"=AVERAGE(R[-{kv.Value}]C:R[-1]C)";

                        }
                        if (i >= 11)
                        {
                            sheetSubTotal.Formula = $"=SUM(R[-{kv.Value}]C:R[-1]C)";
                        }

                        
                    }
                    
                    sheetRowCounts += 1;
                    deptCountDic[key] = deptCountDic[key] + 1;

                    sheetSubTotalRows.Add(sheetRowCounts);        
                }



                for (int i = 0; i < QyxnDataTable.Columns.Count; i++)
                {
                    Range sheetTotal = worksheet.Range.FromLTRB(i, sheetRowCounts, i, sheetRowCounts);
                    switch (i)
                    {
                        case 0:
                            sheetTotal = worksheet.Range.FromLTRB(0, sheetRowCounts - deptCountDic[key], 0, sheetRowCounts );
                            worksheet.MergeCells(sheetTotal);
                            break;
                        case 1:
                            break;
                        case 2:
                            sheetTotal.SetValueFromText("合计：");
                            break;
                        case 3:
                        case 11:
                        case 12:
                            string sumString = "";
                            for (int a = 0; a < sheetSubTotalRows.Count; a++)
                            {
                                sumString = sumString + string.Format($"R{sheetSubTotalRows[a]}C,");
                            }
                            sumString = sumString.TrimEnd(',');

                            sheetTotal.Formula = $"=SUM({sumString})";
                            break;
                        case 4:
                        case 5:
                        case 6:
                        case 7:
                        case 8:
                        case 9:
                        case 10:
                            string avgString = "";
                            for (int a = 0; a < sheetSubTotalRows.Count; a++)
                            {
                                avgString = avgString + string.Format($"R{sheetSubTotalRows[a]}C,");
                            }
                            avgString = avgString.TrimEnd(',');

                            sheetTotal.Formula = $"=AVERAGE({avgString})";
                            break;
                    }
                }

                sheetRowCounts += 1;

            }

            frmMainView.frmMainForm.ssQueryResultView.EndUpdate();
        }
    }

}
