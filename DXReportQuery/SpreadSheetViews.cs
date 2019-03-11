using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data;
using DevExpress.Spreadsheet;
using static DXReportQuery.SpreadSheetStyles;
using System.Drawing;

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

        private static Color GetDeptBackgroundColor(Range cellRange)
        {
            string dept = cellRange.Value.ToString();
            Color color = Color.White;
            switch (dept) {
                case "商超":
                    color = Color.FromArgb(240, 248, 255);
                    break;
                case "生鲜便利":
                    color = Color.FromArgb(255, 228, 196);
                    break;
                case "商锐":
                    color = Color.FromArgb(152, 251, 152);
                    break;
                case "专卖":
                    color = Color.FromArgb(250, 235, 215);
                    break;
                case "孕婴童":
                    color = Color.FromArgb(64, 224, 208);
                    break;
                case "餐饮":
                    color = Color.FromArgb(211, 211, 211);
                    break;
                case "星食客":
                    color = Color.FromArgb(255,182,193);
                    break;
                case "eshop":
                    color = Color.FromArgb(250, 250, 210);
                    break;
                default:
                    color = Color.White;
                    break;
            }
            return color;
        }
        public static void DjwtView()
        {
            Config Config = new Config();
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
            SpreadSheetStyles.SheetStyleInit(workbook);

            workbook.DocumentSettings.R1C1ReferenceStyle = true;

            Worksheet worksheet = SpreadView.GetWorkSheet(workbook, sheetName);
            workbook.Worksheets.ActiveWorksheet = workbook.Worksheets[sheetName];
            worksheet.ActiveView.ShowGridlines = false;
            Range sheetTitleRange = worksheet.Range.FromLTRB(0, sheetRowCounts, 8, sheetRowCounts);
            worksheet.MergeCells(sheetTitleRange);
            sheetTitleRange.Style = workbook.Styles["SheetTitleStyle"];
            sheetTitleRange.SetValueFromText(sheetTitle);
            sheetTitleRange.RowHeight = 30 * 4.16;
            sheetRowCounts += 1;

            List<string> sheetTableHeadList = new List<string> { "行业", "工号", "姓名", "问题登记量", "回访数", "需求登记量", "回访率", "关闭（解决）率", "平均回访周期" };
            for (int i = 0; i < sheetTableHeadList.Count; i++)
            {
                Range sheetTableHeadRange = worksheet.Range.FromLTRB(i, sheetRowCounts, i, sheetRowCounts);
                sheetTableHeadRange.SetValueFromText(sheetTableHeadList[i].ToString());
                sheetTableHeadRange.Style = workbook.Styles["SheetHeadSytle"];
                sheetTableHeadRange.ColumnWidth = 15 * 22;
                if (i >= 6)
                {
                    // sheetTableHeadRange.Style = workbook.Styles["Output"];
                }
            }
            worksheet.Range.FromLTRB(0, sheetRowCounts, 0, sheetRowCounts).RowHeight = 27 * 4.16;
            sheetRowCounts += 1;

            foreach (KeyValuePair<string, int> kv in branchCount)
            {
                DataRow[] dataRows = djwtDataTable.Select($@"ver='{kv.Key}'");
                foreach (DataRow dr in dataRows)
                {
                    for (int i = 0; i < djwtDataTable.Columns.Count; i++)
                    {
                        Range sheetNormal = worksheet.Range.FromLTRB(i, sheetRowCounts, i, sheetRowCounts);
                        sheetNormal.SetValueFromText(dr[i].ToString());
                        sheetNormal.Style = workbook.Styles["SheetNormalSytle"];

                        if (sheetRowCounts % 2 == 0)
                        {
                            sheetNormal.Fill.BackgroundColor = Color.FromArgb(251, 251, 251);
                        }
                        else
                        {
                            sheetNormal.Fill.BackgroundColor = Color.FromArgb(237, 237, 237);
                        }

                        if (i == 6)
                        {
                            sheetNormal.NumberFormat = "0.00%";
                        }
                    }
                    worksheet.Range.FromLTRB(0, sheetRowCounts, 0, sheetRowCounts).RowHeight = 22 * 4.16;
                    
                    sheetRowCounts += 1;
                }

                for (int i = 0; i < djwtDataTable.Columns.Count; i++)
                {
                    Range sheetSubTotal = worksheet.Range.FromLTRB(i, sheetRowCounts, i, sheetRowCounts);
                    sheetSubTotal.RowHeight = 25 * 4.16;
                    sheetSubTotal.Style = workbook.Styles["SheetSubTotalSytle"];
                    if (i == 0)
                    {
                        sheetSubTotal = worksheet.Range.FromLTRB(0, sheetRowCounts - kv.Value, 0, sheetRowCounts);
                        worksheet.MergeCells(sheetSubTotal);
                        sheetSubTotal.Fill.BackgroundColor = GetDeptBackgroundColor(sheetSubTotal);
                    }

                    if (i == 1)
                    {

                        sheetSubTotal.SetValueFromText("小计：");
                        sheetSubTotal.Style = workbook.Styles["SheetSubTotalSytle"];
                    }

                    if (i >= 3 & i <= 6)
                    {
                        sheetSubTotal.Formula = $"=SUM(R[-{kv.Value}]C:R[-1]C)";
                        sheetSubTotal.Style = workbook.Styles["SheetSubTotalSytle"];

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
            Config Config = new Config();
            frmMainView.frmMainForm.ssQueryResultView.BeginUpdate();

            string sheetName = "问题关闭率";
            string sheetTitle1 = string.Format("{0}至{1}", Config.beginTime, Config.endTime);
            string sheetTitle2 = "整体关闭率统计";
            string sheetTitle3 = "vip问题关闭率统计";
            string sheetTitle4 = "处理中问题处理完成情况统计";
            string sheetTitle5 = "问题效能分析报表（全部）";

            int sheetRowCounts = 0;

            DataTable ztgblDataTable = QueryResults.ZtgblQuery();
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
                    if (i == 4 || (i >= 12 & i <= 14))
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


            foreach (DataRow dr in clzWtclDataTable.AsEnumerable())
            {
                for (int i = 0; i < clzWtclDataTable.Columns.Count; i++)
                {
                    Range sheetNormal = worksheet.Range.FromLTRB(i, sheetRowCounts, i, sheetRowCounts);
                    sheetNormal.SetValueFromText(dr[i].ToString());

                    if (i >= 9 & i <= 11)
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
            Config Config = new Config();
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

            List<string> sheetTableHeadList = new List<string> { "行业", "负责人", "省份", "问题取样数", "实际响应速度", "实际解决周期", "实际关闭周期", "有效响应速度", "有效解决周期", "有效关闭周期", "平均回复次数", "响应超时", "处理超时" };
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


                if (deptCountDic[key] > 0)
                {
                    for (int i = 0; i < QyxnDataTable.Columns.Count; i++)
                    {
                        Range sheetTotal = worksheet.Range.FromLTRB(i, sheetRowCounts, i, sheetRowCounts);
                        switch (i)
                        {
                            case 0:
                                sheetTotal = worksheet.Range.FromLTRB(0, sheetRowCounts - deptCountDic[key], 0, sheetRowCounts);
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
                                    sumString = sumString + string.Format($"R[-{sheetRowCounts - sheetSubTotalRows[a] + 1}]C,");
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
                                    avgString = avgString + string.Format($"R[-{sheetRowCounts - sheetSubTotalRows[a] + 1}]C,");
                                }
                                avgString = avgString.TrimEnd(',');

                                sheetTotal.Formula = $"=AVERAGE({avgString})";
                                break;
                        }
                    }
                    sheetRowCounts += 1;
                }
            }
            workbook.DocumentSettings.R1C1ReferenceStyle = false;
            frmMainView.frmMainForm.ssQueryResultView.EndUpdate();
        }

        public static void VIPGblView()
        {
            Config Config = new Config();
            frmMainView.frmMainForm.ssQueryResultView.BeginUpdate();

            string sheetName = "VIP关闭率";
            string sheetTitle = string.Format("{0}至{1}VIP代理商关闭率", Config.beginTime, Config.endTime);
            int sheetRowCounts = 0;             //表单内容当前行数

            IWorkbook workbook = frmMainView.frmMainForm.ssQueryResultView.Document;
            workbook.DocumentSettings.R1C1ReferenceStyle = true;

            Worksheet worksheet = SpreadView.GetWorkSheet(workbook, sheetName);
            workbook.Worksheets.ActiveWorksheet = workbook.Worksheets[sheetName];
            worksheet.ActiveView.ShowGridlines = false;
            Range sheetTitleRange = worksheet.Range.FromLTRB(0, sheetRowCounts, 9, sheetRowCounts);
            worksheet.MergeCells(sheetTitleRange);
            // sheetTitleRange.Style = workbook.Styles["myDjwtSheetTitleStyle"];
            sheetTitleRange.SetValueFromText(sheetTitle);
            sheetRowCounts += 1;

            List<string> sheetTableHeadList = new List<string> { "行业 ", "负责人 ", "公司名称 ", "提交问题数 ", "无修改问题数 ", "负责人处理数量 ", "协助处理数量 ", "问题关闭数量 ", "问题关闭率 ", "负责人处理占比" };
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

            foreach (string key in deptDic.Keys)
            {
                DataTable VIPGblDataTable = QueryResults.VIPGblQuery(key);
                Dictionary<string, int> nameCount = new Dictionary<string, int>();
                List<int> sheetSubTotalRows = new List<int> { };

                deptCountDic.Add(key, VIPGblDataTable.Rows.Count);

                var queryCountResult = from vipGbl in VIPGblDataTable.AsEnumerable()
                                       group vipGbl by new { name = vipGbl.Field<string>("name") }
                                       into g
                                       select new
                                       {
                                           g.Key.name,
                                           count = g.Count()
                                       };

                if (queryCountResult.ToList().Count > 0)
                {
                    queryCountResult.ToList().ForEach(q => nameCount.Add(q.name, q.count));
                }



                foreach (KeyValuePair<string, int> kv in nameCount)
                {
                    DataRow[] dataRows = VIPGblDataTable.Select($@"name='{kv.Key}'");

                    foreach (DataRow dr in dataRows)
                    {
                        for (int i = 0; i < VIPGblDataTable.Columns.Count; i++)
                        {

                            Range sheetNormal = worksheet.Range.FromLTRB(i, sheetRowCounts, i, sheetRowCounts);
                            sheetNormal.SetValueFromText(dr[i].ToString());
                            if (i == 0)
                            {
                                sheetNormal.SetValueFromText(deptDic[key]);
                            }
                            //sheetNormal.Style = workbook.Styles["myDjwtSheetNormalSytle"];
                        }
                        sheetRowCounts += 1;
                    }

                    for (int i = 0; i < VIPGblDataTable.Columns.Count; i++)
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

                        if (i >= 3 && i <= 7)
                        {
                            sheetSubTotal.Formula = $"=SUM(R[-{kv.Value}]C:R[-1]C)";
                        }

                        if (i >= 8 && i <= 9)
                        {
                            sheetSubTotal.Formula = $"=AVERAGE(R[-{kv.Value}]C:R[-1]C)";
                        }

                    }

                    sheetRowCounts += 1;
                    deptCountDic[key] = deptCountDic[key] + 1;

                    sheetSubTotalRows.Add(sheetRowCounts);
                }


                if (deptCountDic[key] > 0)
                {
                    for (int i = 0; i < VIPGblDataTable.Columns.Count; i++)
                    {
                        Range sheetTotal = worksheet.Range.FromLTRB(i, sheetRowCounts, i, sheetRowCounts);
                        switch (i)
                        {
                            case 0:
                                sheetTotal = worksheet.Range.FromLTRB(i, sheetRowCounts - deptCountDic[key], i, sheetRowCounts);
                                worksheet.MergeCells(sheetTotal);
                                break;
                            case 1:
                                break;
                            case 2:
                                sheetTotal.SetValueFromText("合计：");
                                break;
                            case 3:
                            case 4:
                            case 5:
                            case 6:
                            case 7:
                                string sumString = "";
                                for (int a = 0; a < sheetSubTotalRows.Count; a++)
                                {
                                    sumString = sumString + string.Format($"R[-{sheetRowCounts - sheetSubTotalRows[a] + 1}]C,");
                                }
                                sumString = sumString.TrimEnd(',');

                                sheetTotal.Formula = $"=SUM({sumString})";

                                break;
                            case 8:
                            case 9:
                                string avgString = "";
                                for (int a = 0; a < sheetSubTotalRows.Count; a++)
                                {
                                    avgString = avgString + string.Format($"R[-{sheetRowCounts - sheetSubTotalRows[a] + 1}]C,");
                                }
                                avgString = avgString.TrimEnd(',');

                                sheetTotal.Formula = $"=AVERAGE({avgString})";

                                break;

                        }
                    }
                    sheetRowCounts += 1;
                }

            }

            workbook.DocumentSettings.R1C1ReferenceStyle = false;
            frmMainView.frmMainForm.ssQueryResultView.EndUpdate();
        }

        public static void QybbView()
        {
            Config Config = new Config();
            frmMainView.frmMainForm.ssQueryResultView.BeginUpdate();

            string sheetName = "区域报表";
            string sheetTitle = string.Format("{0}至{1}区域报表", Config.beginTime, Config.endTime);
            int sheetRowCounts = 0;             //表单内容当前行数

            IWorkbook workbook = frmMainView.frmMainForm.ssQueryResultView.Document;
            workbook.DocumentSettings.R1C1ReferenceStyle = true;

            Worksheet worksheet = SpreadView.GetWorkSheet(workbook, sheetName);
            workbook.Worksheets.ActiveWorksheet = workbook.Worksheets[sheetName];
            worksheet.ActiveView.ShowGridlines = false;
            Range sheetTitleRange = worksheet.Range.FromLTRB(0, sheetRowCounts, 10, sheetRowCounts);
            worksheet.MergeCells(sheetTitleRange);
            // sheetTitleRange.Style = workbook.Styles["myDjwtSheetTitleStyle"];
            sheetTitleRange.SetValueFromText(sheetTitle);
            sheetRowCounts += 1;

            List<string> sheetTableHeadList = new List<string> { "区域 ","负责人 ","行业 ","省份 ","区域问题数 ","无修改问题数 ","负责人处理数量 ","协助处理数量 ","问题关闭数量 ","区域关闭率 ","负责人处理占比 " };
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

            foreach (string key in deptDic.Keys)
            {
                DataTable QybbDataTable = QueryResults.QybbQuery(key);
                Dictionary<string, int> nameCount = new Dictionary<string, int>();
                List<int> sheetSubTotalRows = new List<int> { };

                deptCountDic.Add(key, QybbDataTable.Rows.Count);

                var queryCountResult = from vipGbl in QybbDataTable.AsEnumerable()
                                       group vipGbl by new { name = vipGbl.Field<string>("name") }
                                       into g
                                       select new
                                       {
                                           g.Key.name,
                                           count = g.Count()
                                       };

                if (queryCountResult.ToList().Count > 0)
                {
                    queryCountResult.ToList().ForEach(q => nameCount.Add(q.name, q.count));
                }



                foreach (KeyValuePair<string, int> kv in nameCount)
                {
                    DataRow[] dataRows = QybbDataTable.Select($@"name='{kv.Key}'");

                    foreach (DataRow dr in dataRows)
                    {
                        for (int i = 0; i < QybbDataTable.Columns.Count; i++)
                        {
                            Range sheetNormal = worksheet.Range.FromLTRB(i+1, sheetRowCounts, i+1, sheetRowCounts);
                            sheetNormal.SetValueFromText(dr[i].ToString());

                            if (i == 0)
                            {
                                sheetNormal = worksheet.Range.FromLTRB(0, sheetRowCounts, 0, sheetRowCounts);
                                sheetNormal.SetValueFromText(deptDic[key]);
                            }
                            //sheetNormal.Style = workbook.Styles["myDjwtSheetNormalSytle"];
                        }
                        sheetRowCounts += 1;
                    }

                    for (int i = 0; i <  QybbDataTable.Columns.Count + 1; i++)
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

                        if (i >= 4 && i <= 8)
                        {
                            sheetSubTotal.Formula = $"=SUM(R[-{kv.Value}]C:R[-1]C)";
                        }

                        if (i >= 9 && i <= 10)
                        {
                            sheetSubTotal.Formula = $"=AVERAGE(R[-{kv.Value}]C:R[-1]C)";
                        }

                    }

                    sheetRowCounts += 1;
                    deptCountDic[key] = deptCountDic[key] + 1;

                    sheetSubTotalRows.Add(sheetRowCounts);
                }


                if (deptCountDic[key] > 0)
                {
                    for (int i = 0; i < QybbDataTable.Columns.Count + 1; i++)
                    {
                        Range sheetTotal = worksheet.Range.FromLTRB(i, sheetRowCounts, i, sheetRowCounts);
                        switch (i)
                        {
                            case 0:
                                sheetTotal = worksheet.Range.FromLTRB(i, sheetRowCounts - deptCountDic[key], i, sheetRowCounts);
                                worksheet.MergeCells(sheetTotal);
                                break;
                            case 1:
                                break;
                            case 2:
                                sheetTotal.SetValueFromText("合计：");
                                break;
                            case 3:
                                break;
                            case 4:
                            case 5:
                            case 6:
                            case 7:
                            case 8:
                                string sumString = "";
                                for (int a = 0; a < sheetSubTotalRows.Count; a++)
                                {
                                    sumString = sumString + string.Format($"R[-{sheetRowCounts - sheetSubTotalRows[a] + 1}]C,");
                                }
                                sumString = sumString.TrimEnd(',');
                                sheetTotal.Formula = $"=SUM({sumString})";
                                break;
                            
                            case 9:
                            case 10:
                                string avgString = "";
                                for (int a = 0; a < sheetSubTotalRows.Count; a++)
                                {
                                    avgString = avgString + string.Format($"R[-{sheetRowCounts - sheetSubTotalRows[a] + 1}]C,");
                                }
                                avgString = avgString.TrimEnd(',');

                                sheetTotal.Formula = $"=AVERAGE({avgString})";

                                break;

                        }
                    }
                    sheetRowCounts += 1;
                }

            }

            workbook.DocumentSettings.R1C1ReferenceStyle = false;
            frmMainView.frmMainForm.ssQueryResultView.EndUpdate();
        }

        public static void GrxnView()
        {
            Config Config = new Config();
            frmMainView.frmMainForm.ssQueryResultView.BeginUpdate();

            string sheetName = "个人效能";
            string sheetTitle = string.Format("{0}至{1}， 实际解决周期商超超过24h/专卖超过30h/餐饮超过22h,用红色区域标记", Config.beginTime, Config.endTime);
            int sheetRowCounts = 0;             //表单内容当前行数

            IWorkbook workbook = frmMainView.frmMainForm.ssQueryResultView.Document;
            workbook.DocumentSettings.R1C1ReferenceStyle = true;

            Worksheet worksheet = SpreadView.GetWorkSheet(workbook, sheetName);
            workbook.Worksheets.ActiveWorksheet = workbook.Worksheets[sheetName];
            worksheet.ActiveView.ShowGridlines = false;
            Range sheetTitleRange = worksheet.Range.FromLTRB(0, sheetRowCounts, 11, sheetRowCounts);
            worksheet.MergeCells(sheetTitleRange);
            // sheetTitleRange.Style = workbook.Styles["myDjwtSheetTitleStyle"];
            sheetTitleRange.SetValueFromText(sheetTitle);
            sheetRowCounts += 1;

            List<string> sheetTableHeadList = new List<string> { "行业","区域负责人","问题取样","实际响应速度","实际解决周期","实际关闭周期","有效响应速度","有效解决周期","有效关闭周期","平均回复次数","响应超时","处理超时"};
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

            foreach (string key in deptDic.Keys)
            {
                DataTable GrxnDataTable = QueryResults.GrxnQuery(key);
                Dictionary<string, int> companyNameCount = new Dictionary<string, int>();

                var queryCountResult = from vipGbl in GrxnDataTable.AsEnumerable()
                                       group vipGbl by new { companyName = vipGbl.Field<string>("companyName") }
                                       into g
                                       select new
                                       {
                                           g.Key.companyName,
                                           count = g.Count()
                                       };

                if (queryCountResult.ToList().Count > 0)
                {
                    queryCountResult.ToList().ForEach(q => companyNameCount.Add(q.companyName, q.count));
                }

                foreach (KeyValuePair<string, int> kv in companyNameCount)
                {
                    DataRow[] dataRows = GrxnDataTable.Select($@"companyName='{kv.Key}'");

                    foreach (DataRow dr in dataRows)
                    {
                        for (int i = 0; i < GrxnDataTable.Columns.Count; i++)
                        {

                            Range sheetNormal = worksheet.Range.FromLTRB(i, sheetRowCounts, i, sheetRowCounts);
                            sheetNormal.SetValueFromText(dr[i].ToString());
                            //sheetNormal.Style = workbook.Styles["myDjwtSheetNormalSytle"];
                        }
                        sheetRowCounts += 1;
                    }

                    for (int i = 0; i < GrxnDataTable.Columns.Count; i++)
                    {
                        Range sheetSubTotal = worksheet.Range.FromLTRB(i, sheetRowCounts, i, sheetRowCounts);

                        if (i == 0)
                        {
                            sheetSubTotal = worksheet.Range.FromLTRB(i, sheetRowCounts - kv.Value, i, sheetRowCounts);
                            worksheet.MergeCells(sheetSubTotal);
                        }

                        if (i == 1)
                        {

                            sheetSubTotal.SetValueFromText("合计：");
                            //sheetSubTotal.Style = workbook.Styles["myDjwtSheetSubTotalSytle"];
                        }

                        if (i == 2 ||( i >= 11 & i <= 12))
                        {
                            sheetSubTotal.Formula = $"=SUM(R[-{kv.Value}]C:R[-1]C)";
                        }

                        if (i >= 3 && i <= 10)
                        {
                            sheetSubTotal.Formula = $"=AVERAGE(R[-{kv.Value}]C:R[-1]C)";
                        }

                    }

                    sheetRowCounts += 1;
                }

            }



            sheetRowCounts += 1;
            string sheetTitle2 = "VIP客户经理个人效能";
            Range sheetTitle2Range = worksheet.Range.FromLTRB(0, sheetRowCounts, 11, sheetRowCounts);
            worksheet.MergeCells(sheetTitle2Range);
            // sheetTitleRange.Style = workbook.Styles["myDjwtSheetTitleStyle"];
            sheetTitle2Range.SetValueFromText(sheetTitle2);
            sheetRowCounts += 1;

            List<string> sheetTableHead2List = new List<string> {"行业","区域负责人","取样问题数量","实际响应速度","实际解决周期","实际关闭周期","有效响应速度","有效解决周期","有效关闭周期","平均回复次数","响应超时","处理超时"};

            for (int i = 0; i < sheetTableHead2List.Count; i++)
            {
                Range sheetTableHead2Range = worksheet.Range.FromLTRB(i, sheetRowCounts, i, sheetRowCounts);
                sheetTableHead2Range.SetValueFromText(sheetTableHead2List[i].ToString());
                //sheetTableHeadRange.Style = workbook.Styles["myDjwtSheetHeadSytle"];

            }
            sheetRowCounts += 1;

            foreach (string key in deptDic.Keys)
            {
                DataTable VIPKhjlGrxnDataTable = QueryResults.VIPKhjlGrxnQuery(key);
                Dictionary<string, int> companyNameCount = new Dictionary<string, int>();

                var queryCountResult = from vipGbl in VIPKhjlGrxnDataTable.AsEnumerable()
                                       group vipGbl by new { companyName = vipGbl.Field<string>("companyName") }
                                       into g
                                       select new
                                       {
                                           g.Key.companyName,
                                           count = g.Count()
                                       };

                if (queryCountResult.ToList().Count > 0)
                {
                    queryCountResult.ToList().ForEach(q => companyNameCount.Add(q.companyName, q.count));
                }

                foreach (KeyValuePair<string, int> kv in companyNameCount)
                {
                    DataRow[] dataRows = VIPKhjlGrxnDataTable.Select($@"companyName='{kv.Key}'");

                    foreach (DataRow dr in dataRows)
                    {
                        for (int i = 0; i < VIPKhjlGrxnDataTable.Columns.Count; i++)
                        {

                            Range sheetNormal = worksheet.Range.FromLTRB(i, sheetRowCounts, i, sheetRowCounts);
                            sheetNormal.SetValueFromText(dr[i].ToString());
                            //sheetNormal.Style = workbook.Styles["myDjwtSheetNormalSytle"];
                        }
                        sheetRowCounts += 1;
                    }

                    for (int i = 0; i < VIPKhjlGrxnDataTable.Columns.Count; i++)
                    {
                        Range sheetSubTotal = worksheet.Range.FromLTRB(i, sheetRowCounts, i, sheetRowCounts);

                        if (i == 0)
                        {
                            sheetSubTotal = worksheet.Range.FromLTRB(i, sheetRowCounts - kv.Value, i, sheetRowCounts);
                            worksheet.MergeCells(sheetSubTotal);
                        }

                        if (i == 1)
                        {

                            sheetSubTotal.SetValueFromText("合计：");
                            //sheetSubTotal.Style = workbook.Styles["myDjwtSheetSubTotalSytle"];
                        }

                        if (i == 2 || (i >= 11 & i <= 12))
                        {
                            sheetSubTotal.Formula = $"=SUM(R[-{kv.Value}]C:R[-1]C)";
                        }

                        if (i >= 3 && i <= 10)
                        {
                            sheetSubTotal.Formula = $"=AVERAGE(R[-{kv.Value}]C:R[-1]C)";
                        }

                    }

                    sheetRowCounts += 1;
                }

            }

            workbook.DocumentSettings.R1C1ReferenceStyle = false;
            frmMainView.frmMainForm.ssQueryResultView.EndUpdate();
        }

        public static void DlsyjView()
        {
            Config Config = new Config();
            frmMainView.frmMainForm.ssQueryResultView.BeginUpdate();

            string sheetName = "代理商预警";
            string sheetTitle = string.Format("{0}至{1}，一周超过5个问题和需求的代理", Config.beginTime, Config.endTime);
            int sheetRowCounts = 0;             //表单内容当前行数
 

            IWorkbook workbook = frmMainView.frmMainForm.ssQueryResultView.Document;

            workbook.DocumentSettings.R1C1ReferenceStyle = true;

            Worksheet worksheet = SpreadView.GetWorkSheet(workbook, sheetName);
            workbook.Worksheets.ActiveWorksheet = workbook.Worksheets[sheetName];
            worksheet.ActiveView.ShowGridlines = false;
            Range sheetTitleRange = worksheet.Range.FromLTRB(0, sheetRowCounts, 11, sheetRowCounts);
            worksheet.MergeCells(sheetTitleRange);
            // sheetTitleRange.Style = workbook.Styles["myDjwtSheetTitleStyle"];
            sheetTitleRange.SetValueFromText(sheetTitle);
            sheetRowCounts += 1;

            List<string> sheetTableHeadList = new List<string> { "行业", "代理商", "版本", "负责人", "问题次数", "联系人", "联系电话", "是否同一客户", "是否有待跟进问题，编号多少", "目前急需解决的问题", "回访信息", "周五回访" };
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

            foreach (string key in deptDic.Keys)
            {
                DataTable dlsyjDataTable = QueryResults.DlsyjQuery(key);
                Dictionary<string, int> companyNameCount = new Dictionary<string, int>();

                var queryCountResult = from dlsyj in dlsyjDataTable.AsEnumerable()
                                       group dlsyj by new { companyName = dlsyj.Field<string>("companyName") }
                                       into g
                                       select new
                                       {
                                           g.Key.companyName,
                                           count = g.Count()
                                       };

                if (queryCountResult.ToList().Count > 0)
                {
                    queryCountResult.ToList().ForEach(q => companyNameCount.Add(q.companyName, q.count));
                }

                foreach (KeyValuePair<string, int> kv in companyNameCount)
                {
                    DataRow[] dataRows = dlsyjDataTable.Select($@"companyName='{kv.Key}'");

                    foreach (DataRow dr in dataRows)
                    {
                        for (int i = 0; i < dlsyjDataTable.Columns.Count; i++)
                        {

                            Range sheetNormal = worksheet.Range.FromLTRB(i, sheetRowCounts, i, sheetRowCounts);
                            sheetNormal.SetValueFromText(dr[i].ToString());
                            //sheetNormal.Style = workbook.Styles["myDjwtSheetNormalSytle"];
                        }
                        sheetRowCounts += 1;
                    }

                    for (int i = 0; i < dlsyjDataTable.Columns.Count; i++)
                    {
                        Range sheetSubTotal = worksheet.Range.FromLTRB(i, sheetRowCounts, i, sheetRowCounts);
                        if (i == 0)
                        {
                            sheetSubTotal = worksheet.Range.FromLTRB(0, sheetRowCounts - kv.Value , 0, sheetRowCounts - 1);
                            worksheet.MergeCells(sheetSubTotal);
                        }
                    }
                }
            }


            workbook.DocumentSettings.R1C1ReferenceStyle = false;
            frmMainView.frmMainForm.ssQueryResultView.EndUpdate();
        }

        public static void VIPDlsyjView()
        {
            Config Config = new Config();
            frmMainView.frmMainForm.ssQueryResultView.BeginUpdate();

            string sheetName = "VIP代理商预警";
            string sheetTitle = string.Format("{0}至{1}，一周超过5个问题和需求的代理", Config.beginTime, Config.endTime);
            int sheetRowCounts = 0;             //表单内容当前行数


            IWorkbook workbook = frmMainView.frmMainForm.ssQueryResultView.Document;

            workbook.DocumentSettings.R1C1ReferenceStyle = true;

            Worksheet worksheet = SpreadView.GetWorkSheet(workbook, sheetName);
            workbook.Worksheets.ActiveWorksheet = workbook.Worksheets[sheetName];
            worksheet.ActiveView.ShowGridlines = false;
            Range sheetTitleRange = worksheet.Range.FromLTRB(0, sheetRowCounts, 11, sheetRowCounts);
            worksheet.MergeCells(sheetTitleRange);
            // sheetTitleRange.Style = workbook.Styles["myDjwtSheetTitleStyle"];
            sheetTitleRange.SetValueFromText(sheetTitle);
            sheetRowCounts += 1;

            List<string> sheetTableHeadList = new List<string> { "行业", "代理商", "版本", "负责人", "问题次数", "联系人", "联系电话", "是否同一客户", "是否有待跟进问题，编号多少", "目前急需解决的问题", "回访信息", "周五回访" };
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

            foreach (string key in deptDic.Keys)
            {
                DataTable VIPDlsyjDataTable = QueryResults.VIPDlsyjQuery(key);
                Dictionary<string, int> deptCount = new Dictionary<string, int>();

                var queryCountResult = from dlsyj in VIPDlsyjDataTable.AsEnumerable()
                                       group dlsyj by new { dept = dlsyj.Field<string>("dept") }
                                       into g
                                       select new
                                       {
                                           g.Key.dept,
                                           count = g.Count()
                                       };

                if (queryCountResult.ToList().Count > 0)
                {
                    queryCountResult.ToList().ForEach(q => deptCount.Add(q.dept, q.count));
                }

                foreach (KeyValuePair<string, int> kv in deptCount)
                {
                    DataRow[] dataRows = VIPDlsyjDataTable.Select($@"dept='{kv.Key}'");

                    foreach (DataRow dr in dataRows)
                    {
                        for (int i = 0; i < VIPDlsyjDataTable.Columns.Count; i++)
                        {

                            Range sheetNormal = worksheet.Range.FromLTRB(i, sheetRowCounts, i, sheetRowCounts);
                            sheetNormal.SetValueFromText(dr[i].ToString());
                            //sheetNormal.Style = workbook.Styles["myDjwtSheetNormalSytle"];
                        }
                        sheetRowCounts += 1;
                    }

                    for (int i = 0; i < VIPDlsyjDataTable.Columns.Count; i++)
                    {
                        Range sheetSubTotal = worksheet.Range.FromLTRB(i, sheetRowCounts, i, sheetRowCounts);
                        if (i == 0)
                        {
                            sheetSubTotal = worksheet.Range.FromLTRB(0, sheetRowCounts - kv.Value, 0, sheetRowCounts - 1);
                            worksheet.MergeCells(sheetSubTotal);
                        }
                    }
                }
            }


            workbook.DocumentSettings.R1C1ReferenceStyle = false;
            frmMainView.frmMainForm.ssQueryResultView.EndUpdate();
        }

        public static void WtyjView()
        {
            Config Config = new Config();
            frmMainView.frmMainForm.ssQueryResultView.BeginUpdate();

            string sheetName = "问题预警";
            string sheetTitle = string.Format("{0}至{1}，问题预警", Config.beginTime, Config.endTime);
            int sheetRowCounts = 0;             //表单内容当前行数

            DataTable wtyjDataTable = QueryResults.WtyjQuery();
            Dictionary<string, int> branchCount = new Dictionary<string, int>();

            var queryCountResult = from wtyj in wtyjDataTable.AsEnumerable()
                                   group wtyj by new { version = wtyj.Field<string>("version") }
                                   into g
                                   select new
                                   {
                                       g.Key.version,
                                       count = g.Count()
                                   };

            if (queryCountResult.ToList().Count > 0)
            {
                queryCountResult.ToList().ForEach(q => branchCount.Add(q.version, q.count));
            }

            IWorkbook workbook = frmMainView.frmMainForm.ssQueryResultView.Document;

            workbook.DocumentSettings.R1C1ReferenceStyle = true;

            Worksheet worksheet = SpreadView.GetWorkSheet(workbook, sheetName);
            workbook.Worksheets.ActiveWorksheet = workbook.Worksheets[sheetName];
            worksheet.ActiveView.ShowGridlines = false;
            Range sheetTitleRange = worksheet.Range.FromLTRB(0, sheetRowCounts, 16, sheetRowCounts);
            worksheet.MergeCells(sheetTitleRange);
            // sheetTitleRange.Style = workbook.Styles["myDjwtSheetTitleStyle"];
            sheetTitleRange.SetValueFromText(sheetTitle);
            sheetRowCounts += 1;

            List<string> sheetTableHeadList = new List<string> { "行业 ", "版本 ", "负责人 ", "问题编号 ", "省份 ", "主状态 ", "研发跟进 ", "首次提交时间 ", "研发处理时长(小时) ", "最后回复时间 ", "解决时间（小时） ", "联系人 ", " ", "联系电话 ", "问题处理进度及情况 ", "问题未关闭原因 ", "回访信息 ", "周五回访" };
            for (int i = 0; i < sheetTableHeadList.Count; i++)
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
                DataRow[] dataRows = wtyjDataTable.Select($@"version='{kv.Key}'");
                foreach (DataRow dr in dataRows)
                {
                    for (int i = 0; i < wtyjDataTable.Columns.Count; i++)
                    {
                        Range sheetNormal = worksheet.Range.FromLTRB(i, sheetRowCounts, i, sheetRowCounts);
                        sheetNormal.SetValueFromText(dr[i].ToString());
                        //sheetNormal.Style = workbook.Styles["myDjwtSheetNormalSytle"];
                    }
                    sheetRowCounts += 1;
                }

                for (int i = 0; i < wtyjDataTable.Columns.Count; i++)
                {
                    Range sheetSubTotal = worksheet.Range.FromLTRB(i, sheetRowCounts, i, sheetRowCounts);

                    if (i == 0)
                    {
                        sheetSubTotal = worksheet.Range.FromLTRB(0, sheetRowCounts - kv.Value, 0, sheetRowCounts - 1);
                        worksheet.MergeCells(sheetSubTotal);
                    }
                }
            }

            workbook.DocumentSettings.R1C1ReferenceStyle = false;
            frmMainView.frmMainForm.ssQueryResultView.EndUpdate();
        }

        public static void WtxqzbView()
        {

        }

        public static void ZzsktjView()
        {
            Config Config = new Config();
            frmMainView.frmMainForm.ssQueryResultView.BeginUpdate();

            string sheetName = "周转知识库";
            string sheetTitle = string.Format("{0}至{1}转知识库统计（二线）", Config.beginTime, Config.endTime);
            int sheetRowCounts = 0;             //表单内容当前行数

            DataTable ZzsktjDataTable = QueryResults.ZzsktjQuery();
            Dictionary<string, int> branchCount = new Dictionary<string, int>();

            var queryCountResult = from djwt in ZzsktjDataTable.AsEnumerable()
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

            workbook.DocumentSettings.R1C1ReferenceStyle = true;

            Worksheet worksheet = SpreadView.GetWorkSheet(workbook, sheetName);
            workbook.Worksheets.ActiveWorksheet = workbook.Worksheets[sheetName];
            worksheet.ActiveView.ShowGridlines = false;
            Range sheetTitleRange = worksheet.Range.FromLTRB(0, sheetRowCounts, 2, sheetRowCounts);
            worksheet.MergeCells(sheetTitleRange);
            sheetTitleRange.Style = workbook.Styles["SheetTitleStyle"];
            sheetTitleRange.SetValueFromText(sheetTitle);
            sheetRowCounts += 1;

            List<string> sheetTableHeadList = new List<string> { "行业", "处理人", "转知识库数量" };
            for (int i = 0; i < sheetTableHeadList.Count; i++)
            {
                Range sheetTableHeadRange = worksheet.Range.FromLTRB(i, sheetRowCounts, i, sheetRowCounts);
                sheetTableHeadRange.SetValueFromText(sheetTableHeadList[i].ToString());
                sheetTableHeadRange.Style = workbook.Styles["SheetHeadSytle"];
                if (i >= 6)
                {
                    // sheetTableHeadRange.Style = workbook.Styles["Output"];
                }
            }
            sheetRowCounts += 1;

            foreach (KeyValuePair<string, int> kv in branchCount)
            {
                DataRow[] dataRows = ZzsktjDataTable.Select($@"ver='{kv.Key}'");
                foreach (DataRow dr in dataRows)
                {
                    for (int i = 0; i < ZzsktjDataTable.Columns.Count; i++)
                    {
                        Range sheetNormal = worksheet.Range.FromLTRB(i, sheetRowCounts, i, sheetRowCounts);
                        sheetNormal.SetValueFromText(dr[i].ToString());
                        sheetNormal.Style = workbook.Styles["SheetNormalSytle"];
                    }
                    sheetRowCounts += 1;
                }

                for (int i = 0; i < ZzsktjDataTable.Columns.Count; i++)
                {
                    Range sheetSubTotal = worksheet.Range.FromLTRB(i, sheetRowCounts, i, sheetRowCounts);

                    if (i == 0)
                    {
                        sheetSubTotal = worksheet.Range.FromLTRB(0, sheetRowCounts - kv.Value, 0, sheetRowCounts -1);
                        worksheet.MergeCells(sheetSubTotal);
                    }
                }
            }



            //知识库处理数量统计报表开始
            string sheetTitle2 = string.Format("{0}至{1}知识库处理数量统计", Config.beginTime, Config.endTime);
            sheetRowCounts = 0;             //表单内容行数

            Range sheetTitle2Range = worksheet.Range.FromLTRB(0+5, sheetRowCounts, 3+5, sheetRowCounts);
            worksheet.MergeCells(sheetTitle2Range);
            // sheetTitleRange.Style = workbook.Styles["myDjwtSheetTitleStyle"];
            sheetTitle2Range.SetValueFromText(sheetTitle2);
            sheetRowCounts += 1;

            List<string> sheetTableHead2List = new List<string> { "姓名", "知识库", "不处理", "合计" };
            for (int i = 0; i < sheetTableHead2List.Count; i++)
            {
                Range sheetTableHead2Range = worksheet.Range.FromLTRB(i+5, sheetRowCounts, i+5, sheetRowCounts);
                sheetTableHead2Range.SetValueFromText(sheetTableHead2List[i].ToString());
                //sheetTableHeadRange.Style = workbook.Styles["myDjwtSheetHeadSytle"];
                if (i >= 6)
                {
                    // sheetTableHeadRange.Style = workbook.Styles["Output"];
                }
            }
            sheetRowCounts += 1;

            DataTable ZskclsltjDataTable = QueryResults.ZskclsltjQuery();
            DataRow[] zskclsltjDataRows = ZskclsltjDataTable.Select();

            foreach (DataRow dr in zskclsltjDataRows)
            {
                for (int i = 0; i < ZskclsltjDataTable.Columns.Count; i++)
                {
                    Range sheetNormal = worksheet.Range.FromLTRB(i+5, sheetRowCounts, i+5, sheetRowCounts);
                    sheetNormal.SetValueFromText(dr[i].ToString());
                    //sheetNormal.Style = workbook.Styles["myDjwtSheetNormalSytle"];
                }
                sheetRowCounts += 1;
            }

            sheetRowCounts += 2;

            //知识库整理统计
            string sheetTitle3 = string.Format("{0}至{1}知识库整理统计", Config.beginTime, Config.endTime);
           
            Range sheetTitle3Range = worksheet.Range.FromLTRB(0 + 5, sheetRowCounts, 3 + 5, sheetRowCounts);
            worksheet.MergeCells(sheetTitle3Range);
            // sheetTitleRange.Style = workbook.Styles["myDjwtSheetTitleStyle"];
            sheetTitle3Range.SetValueFromText(sheetTitle3);
            sheetRowCounts += 1;

            List<string> sheetTableHead3List = new List<string> { "项目", "整理人", "整理数量", "新增数量" };
            for (int i = 0; i < sheetTableHead2List.Count; i++)
            {
                Range sheetTableHead3Range = worksheet.Range.FromLTRB(i + 5, sheetRowCounts, i + 5, sheetRowCounts);
                sheetTableHead3Range.SetValueFromText(sheetTableHead3List[i].ToString());
                //sheetTableHeadRange.Style = workbook.Styles["myDjwtSheetHeadSytle"];
                if (i >= 6)
                {
                    // sheetTableHeadRange.Style = workbook.Styles["Output"];
                }
            }
            sheetRowCounts += 1;

            DataTable ZskzltjDataTable = QueryResults.ZskzltjQuery();
            DataRow[] zskzltjDataRows = ZskzltjDataTable.Select();

            foreach (DataRow dr in zskzltjDataRows)
            {
                for (int i = 0; i < ZskzltjDataTable.Columns.Count; i++)
                {
                    Range sheetNormal = worksheet.Range.FromLTRB(i + 5, sheetRowCounts, i + 5, sheetRowCounts);
                    sheetNormal.SetValueFromText(dr[i].ToString());
                    //sheetNormal.Style = workbook.Styles["myDjwtSheetNormalSytle"];
                }
                sheetRowCounts += 1;
            }

            for (int i = 0; i < ZskzltjDataTable.Columns.Count; i++)
            {
                if (i == 1)
                {
                    Range sheetSubTotal = worksheet.Range.FromLTRB(i + 5, sheetRowCounts, i + 5, sheetRowCounts);
                    sheetSubTotal.SetValueFromText("小计");
                }

                if (i >= 2 )
                {
                    Range sheetSubTotal = worksheet.Range.FromLTRB(i + 5, sheetRowCounts, i + 5, sheetRowCounts);
                    sheetSubTotal.Formula = $"=SUM(R[-{ZskzltjDataTable.Columns.Count - 1}]C:R[-1]C)";
                }

            }


            workbook.DocumentSettings.R1C1ReferenceStyle = false;
            frmMainView.frmMainForm.ssQueryResultView.EndUpdate();
        }

    }

}
