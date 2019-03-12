using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DevExpress.Spreadsheet;
using System.Drawing;

namespace DXReportQuery
{
    public static class SpreadSheetStyles
    {
        private static bool SheetStyleInitFlag = false;

        public static void SheetStyleInit(IWorkbook workBook)
        {
            if (!SheetStyleInitFlag)
            {
                SheetStyle.SheetTitleStyle(workBook);
                SheetStyle.SheetHeadStyle(workBook);
                SheetStyle.SheetNormalStyle(workBook);
                SheetStyle.SheetSubTotalStyle(workBook);
                SheetStyle.SheetSumTotalStyle(workBook);
            }

            SheetStyleInitFlag = true;
        }
        public class SheetStyle
        {
            public static void SheetTitleStyle(IWorkbook workBook)
            {
                Style SheetTitleStyle = workBook.Styles.Add("SheetTitleStyle");
                //SheetTitleStyle.CopyFrom(BuiltInStyleId.Heading2);
                SheetTitleStyle.BeginUpdate();
                try
                {
                    SheetTitleStyle.Font.Name = "宋体";
                    SheetTitleStyle.Font.Size = 18;
                    SheetTitleStyle.Font.Color = Color.FromArgb(52, 150, 151);
                    SheetTitleStyle.Alignment.Horizontal = SpreadsheetHorizontalAlignment.Left;
                    SheetTitleStyle.Alignment.Vertical = SpreadsheetVerticalAlignment.Center;
                }
                finally
                {
                    SheetTitleStyle.EndUpdate();
                }
            }

            public static void SheetHeadStyle(IWorkbook workBook)
            {
                Style SheetHeadStyle = workBook.Styles.Add("SheetHeadStyle");
                //SheetHeadStyle.CopyFrom(BuiltInStyleId.Accent6);

                SheetHeadStyle.BeginUpdate();
                try
                {
                    
                    SheetHeadStyle.Borders.TopBorder.Color = Color.FromArgb(166,166,166);
                    SheetHeadStyle.Borders.BottomBorder.Color = Color.FromArgb(166, 166, 166);

                    SheetHeadStyle.Borders.LeftBorder.Color = Color.White;
                    SheetHeadStyle.Borders.RightBorder.Color = Color.White;

                    SheetHeadStyle.Fill.BackgroundColor = Color.FromArgb(242,242,242);
                    SheetHeadStyle.Alignment.Horizontal = SpreadsheetHorizontalAlignment.Center;
                    SheetHeadStyle.Alignment.Vertical = SpreadsheetVerticalAlignment.Center;

                    SheetHeadStyle.Font.Size = 11;
                    SheetHeadStyle.Font.Color = Color.FromArgb(52, 150, 151);

                }           
                finally
                {
                    SheetHeadStyle.EndUpdate();
                }
            }

            public static void SheetNormalStyle(IWorkbook workBook)
            {
                Style SheetNormalStyle = workBook.Styles.Add("SheetNormalStyle");
                SheetNormalStyle.CopyFrom(BuiltInStyleId.Normal);

                SheetNormalStyle.BeginUpdate();
                try
                {
                    SheetNormalStyle.Font.Size = 11;
                    SheetNormalStyle.Alignment.Horizontal = SpreadsheetHorizontalAlignment.Center;
                    SheetNormalStyle.Alignment.Vertical = SpreadsheetVerticalAlignment.Center;
                }
                finally
                {
                    SheetNormalStyle.EndUpdate();
                }
            }

            public static void SheetSubTotalStyle(IWorkbook workBook)
            {
                Style SheetSubTotalStyle = workBook.Styles.Add("SheetSubTotalStyle");
                SheetSubTotalStyle.CopyFrom(BuiltInStyleId.Neutral);

                SheetSubTotalStyle.BeginUpdate();
                try
                {
                    SheetSubTotalStyle.Font.Size = 11;
                    SheetSubTotalStyle.Font.Bold = true;
                    SheetSubTotalStyle.Alignment.Horizontal = SpreadsheetHorizontalAlignment.Center;
                    SheetSubTotalStyle.Alignment.Vertical = SpreadsheetVerticalAlignment.Center;

                }
                finally
                {
                    SheetSubTotalStyle.EndUpdate();
                }
            }

            public static void SheetSumTotalStyle(IWorkbook workBook)
            {
                Style SheetSumTotalStyle = workBook.Styles.Add("SheetSumTotalStyle");
                SheetSumTotalStyle.CopyFrom(BuiltInStyleId.Input);

                SheetSumTotalStyle.BeginUpdate();
                try
                {
                    SheetSumTotalStyle.Font.Size = 12;
                    SheetSumTotalStyle.Font.Bold = true;
                    SheetSumTotalStyle.Borders.SetAllBorders(Color.White, BorderLineStyle.None);
                    SheetSumTotalStyle.Alignment.Horizontal = SpreadsheetHorizontalAlignment.Center;
                    SheetSumTotalStyle.Alignment.Vertical = SpreadsheetVerticalAlignment.Center;

                }
                finally
                {
                    SheetSumTotalStyle.EndUpdate();
                }
            }

           
        }





}
}
