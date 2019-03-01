using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
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
            var fluent = mvvmContext1.OfType<frmMainViewModel>();
            fluent.SetBinding(beiStartDate, e => e.EditValue, x => x.startDate);
            fluent.SetBinding(beiEndDate, e => e.EditValue, x => x.endDate);

            beiStartDate.EditValueChanged += (s, e) => fluent.ViewModel.SetEndDate();

            nbcRcgzbb.LinkClicked += (s, e) => fluent.ViewModel.nbcRcgzbb_LinkClicked(s, e);
        }
    }
}