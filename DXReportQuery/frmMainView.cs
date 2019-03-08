using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace DXReportQuery
{
    public partial class frmMainView : DevExpress.XtraBars.Ribbon.RibbonForm
    {
        public static frmMainView frmMainForm;
        public frmMainView()
        {
            InitializeComponent();
            if (!mvvmContext1.IsDesignMode)
                InitializeBindings();
            frmMainForm = this;
        }

        void InitializeBindings()
        {
            string beginDate = DateTime.Now.AddDays(Convert.ToDouble((0 - Convert.ToInt16(DateTime.Now.DayOfWeek))) - 14 + 6).ToString("yyyy-MM-dd");
            DateTimeFormatInfo dtFormat = new DateTimeFormatInfo();
            dtFormat.ShortDatePattern = "yyyy-MM-dd";
            string endDate = Convert.ToDateTime(beginDate, dtFormat).AddDays(6).ToString("yyyy-MM-dd");
            
            var frmMainViewModel = new frmMainViewModel(beginDate, endDate);
            mvvmContext1.SetViewModel(typeof(frmMainViewModel), frmMainViewModel);

            var fluent = mvvmContext1.OfType<frmMainViewModel>();
            fluent.SetBinding(beiStartDate, e => e.EditValue, x => x.StartDate);
            fluent.SetBinding(beiEndDate, e => e.EditValue, x => x.EndDate);

            beiStartDate.EditValueChanged += (s, e) => fluent.ViewModel.SetStartDate();
            beiStartDate.EditValueChanged += (s, e) => fluent.ViewModel.SetEndDate();
            beiEndDate.EditValueChanged += (s, e) => fluent.ViewModel.SetEndDate2();

            nbcRcgzbb.LinkClicked += (s, e) => fluent.ViewModel.nbcRcgzbb_LinkClicked(s, e);
        }

      
    }
}