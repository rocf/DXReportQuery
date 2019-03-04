using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DevExpress.Spreadsheet;

namespace DXReportQuery
{
    public static class SpreadSheetStyles
    {
        private static bool DjwtSheetStyleInitFlag = false;

        public static void DjwtSheetStyleInit(IWorkbook workBook)
        {
            if (!DjwtSheetStyleInitFlag)
            {
                DjwtSheetStyle.DjwtSheetTitleStyle(workBook);
                DjwtSheetStyle.DjwtSheetHeadSytle(workBook);
                DjwtSheetStyle.DjwtSheetNormalSytle(workBook);
                DjwtSheetStyle.DjwtSheetSubTotalSytle(workBook);
            }

            DjwtSheetStyleInitFlag = true;
        }
        public class DjwtSheetStyle
        {
            public static void DjwtSheetTitleStyle(IWorkbook workBook)
            {
                Style myDjwtSheetTitleStyle = workBook.Styles.Add("myDjwtSheetTitleStyle");
                myDjwtSheetTitleStyle.CopyFrom(BuiltInStyleId.Title);
                myDjwtSheetTitleStyle.BeginUpdate();
                try
                {

                }
                finally
                {
                    myDjwtSheetTitleStyle.EndUpdate();
                }
            }

            public static void DjwtSheetHeadSytle(IWorkbook workBook)
            {
                Style myDjwtSheetHeadSytle = workBook.Styles.Add("myDjwtSheetHeadSytle");
                myDjwtSheetHeadSytle.CopyFrom(BuiltInStyleId.Accent6);

                myDjwtSheetHeadSytle.BeginUpdate();
                try
                {

                }
                finally
                {
                    myDjwtSheetHeadSytle.EndUpdate();
                }
            }

            public static void DjwtSheetNormalSytle(IWorkbook workBook)
            {
                Style myDjwtSheetNormalSytle = workBook.Styles.Add("myDjwtSheetNormalSytle");
                myDjwtSheetNormalSytle.CopyFrom(BuiltInStyleId.Normal);

                myDjwtSheetNormalSytle.BeginUpdate();
                try
                {

                }
                finally
                {
                    myDjwtSheetNormalSytle.EndUpdate();
                }
            }

            public static void DjwtSheetSubTotalSytle(IWorkbook workBook)
            {
                Style myDjwtSheetSubTotalSytle = workBook.Styles.Add("myDjwtSheetSubTotalSytle");
                myDjwtSheetSubTotalSytle.CopyFrom(BuiltInStyleId.Neutral);

                myDjwtSheetSubTotalSytle.BeginUpdate();
                try
                {

                }
                finally
                {
                    myDjwtSheetSubTotalSytle.EndUpdate();
                }
            }
        }





}
}
