using DevExpress.Mvvm.DataAnnotations;
using DevExpress.XtraNavBar;
using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace DXReportQuery
{
    [POCOViewModel()]
    public class frmMainViewModel
    {
        public frmMainViewModel()
        {
            SetStartDate();
            SetEndDate();
        }

        public virtual string startDate { get; set; }
        public virtual string endDate { get; set; }

        public static string beginDate { get; set; }
        public static string lastDate { get; set; }

        public string GetStartDate()
        {
            return DateTime.Now.AddDays(Convert.ToDouble((0 - Convert.ToInt16(DateTime.Now.DayOfWeek))) - 14 + 6).ToString("yyyy-MM-dd");
        }
        public void SetStartDate()
        {
            frmMainViewModel.beginDate = startDate = GetStartDate();
        }
        public string GetEndDate()
        {
            DateTimeFormatInfo dtFormat = new DateTimeFormatInfo();
            dtFormat.ShortDatePattern = "yyyy-MM-dd";
            return Convert.ToDateTime(startDate, dtFormat).AddDays(6).ToString("yyyy-MM-dd");
        }
        public void SetEndDate()
        {
            frmMainViewModel.lastDate = endDate = GetEndDate();
        }
        internal void nbcRcgzbb_LinkClicked(object sender, NavBarLinkEventArgs e)
        {
            /*
            
            */
            switch(e.Link.ItemName)
            {
                case "nbiDjwt":
                    SpreadView.DjwtView();
                    break;
                case "nbiWtgbl":
                    SpreadView.WtgblView();
                    break;
                default:
                    break;
            }
        }

    }
}
