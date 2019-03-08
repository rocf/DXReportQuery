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

        string startDate { get; set; }
        string endDate { get; set; }
        public frmMainViewModel(string sDate, string eDate)
        {
           beginDate =  this.startDate = sDate;
           lastDate = this.endDate = eDate;
        }

        public virtual string StartDate
        {
            get { return startDate; }
            set
            {
                if (startDate == value) return;
                startDate = value;
                OnStartDateChanged();
            }
        }

        public virtual string EndDate
        {
            get { return endDate; }
            set
            {
                if (endDate == value) return;
                endDate = value;
                OnEndDateChanged();
            }
        }
        void OnStartDateChanged()
        {
            EventHandler h = StartDateChanged;
            if (h != null) h(this, EventArgs.Empty);
        }

        void OnEndDateChanged()
        {
            EventHandler h = EndDateChanged;
            if (h != null) h(this, EventArgs.Empty);
        }

        public event EventHandler StartDateChanged;
        public event EventHandler EndDateChanged;


        public static string beginDate { get; set; }
        public static string lastDate { get; set; }

        public string GetStartDate()
        {
            DateTimeFormatInfo dtFormat = new DateTimeFormatInfo();
            dtFormat.ShortDatePattern = "yyyy-MM-dd";
            return Convert.ToDateTime(frmMainView.frmMainForm.beiStartDate.EditValue.ToString(), dtFormat).ToString("yyyy-MM-dd");
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

        public void SetEndDate2()
        {
            DateTimeFormatInfo dtFormat = new DateTimeFormatInfo();
            dtFormat.ShortDatePattern = "yyyy-MM-dd";
            frmMainViewModel.lastDate = endDate = Convert.ToDateTime(frmMainView.frmMainForm.beiEndDate.EditValue.ToString(), dtFormat).ToString("yyyy-MM-dd");

        }
        internal void nbcRcgzbb_LinkClicked(object sender, NavBarLinkEventArgs e)
        {
            /*
            
            */
            switch(e.Link.ItemName)
            {
                case "nbiDjwt":
                    frmMainView.frmMainForm.splashScreenManager2.ShowWaitForm();
                    SpreadView.DjwtView();
                    frmMainView.frmMainForm.splashScreenManager2.CloseWaitForm();
                    break;
                case "nbiWtgbl":
                    frmMainView.frmMainForm.splashScreenManager2.ShowWaitForm();
                    SpreadView.WtgblView();
                    frmMainView.frmMainForm.splashScreenManager2.CloseWaitForm();
                    break;
                case "nbiWtmxgbl":
                    break;
                case "nbiQyxn":
                    frmMainView.frmMainForm.splashScreenManager2.ShowWaitForm();
                    SpreadView.QyxnView();
                    frmMainView.frmMainForm.splashScreenManager2.CloseWaitForm();
                    break;
                case "nbiVIPgbl":
                    frmMainView.frmMainForm.splashScreenManager2.ShowWaitForm();
                    SpreadView.VIPGblView();
                    frmMainView.frmMainForm.splashScreenManager2.CloseWaitForm();
                    break;
                case "nbiQybb":
                    frmMainView.frmMainForm.splashScreenManager2.ShowWaitForm();
                    SpreadView.QybbView();
                    frmMainView.frmMainForm.splashScreenManager2.CloseWaitForm();
                    break;
                case "nbiGrxn":
                    frmMainView.frmMainForm.splashScreenManager2.ShowWaitForm();
                    SpreadView.GrxnView();
                    frmMainView.frmMainForm.splashScreenManager2.CloseWaitForm();
                    break;
                case "nbiDlsyj":
                    frmMainView.frmMainForm.splashScreenManager2.ShowWaitForm();
                    SpreadView.DlsyjView();
                    frmMainView.frmMainForm.splashScreenManager2.CloseWaitForm();
                    break;
                case "nbiVIPdlsyj":
                    frmMainView.frmMainForm.splashScreenManager2.ShowWaitForm();
                    SpreadView.VIPDlsyjView();
                    frmMainView.frmMainForm.splashScreenManager2.CloseWaitForm();
                    break;
                case "nbiWtyj":
                    frmMainView.frmMainForm.splashScreenManager2.ShowWaitForm();
                    SpreadView.WtyjView();
                    frmMainView.frmMainForm.splashScreenManager2.CloseWaitForm();
                    break;
                case "nbiWtxqzbl":
                    frmMainView.frmMainForm.splashScreenManager2.ShowWaitForm();
                    SpreadView.WtxqzbView();
                    frmMainView.frmMainForm.splashScreenManager2.CloseWaitForm();
                    break;
                case "nbiZzzsk":
                    frmMainView.frmMainForm.splashScreenManager2.ShowWaitForm();
                    SpreadView.ZzsktjView();
                    frmMainView.frmMainForm.splashScreenManager2.CloseWaitForm();
                    break;
                default:
                    break;
            }
        }

    }
}
