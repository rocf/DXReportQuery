using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DevExpress.Spreadsheet;

namespace DXReportQuery
{
   class SpreadSheetStyles
    {
        public class DjwtSheetStyle
        {
            public static Style DjwtSheetTitleStyle(IWorkbook workBook)
            {
                Style myDjwtSheetHeadStyle = workBook.Styles.Add("myDjwtSheetHeadStyle");
                myDjwtSheetHeadStyle.CopyFrom(BuiltInStyleId.Title);
                myDjwtSheetHeadStyle.BeginUpdate();
                try
                {

                }
                finally
                {
                    myDjwtSheetHeadStyle.EndUpdate();
                }
                return myDjwtSheetHeadStyle;
            }
        }





}
}
