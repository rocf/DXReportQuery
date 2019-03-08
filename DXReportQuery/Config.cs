using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;

namespace DXReportQuery
{
    class Config
    {

        public Config()
        {
            beginTime = GetStartDate();
            endTime = GetEndDate();
        }
        internal string beginTime;
        internal string endTime;

        internal string connectionString = "Data Source=120.78.131.78,6985;Initial Catalog=iss_support;User Id=iss;Password=support@siss2014";

        public string GetStartDate()
        {
            DateTimeFormatInfo dtFormat = new DateTimeFormatInfo();
            dtFormat.ShortDatePattern = "yyyy-MM-dd";
            return Convert.ToDateTime(frmMainView.frmMainForm.beiStartDate.EditValue.ToString(), dtFormat).ToString("yyyy-MM-dd");
        }

        public string GetEndDate()
        {
            DateTimeFormatInfo dtFormat = new DateTimeFormatInfo();
            dtFormat.ShortDatePattern = "yyyy-MM-dd";
            return Convert.ToDateTime(frmMainView.frmMainForm.beiEndDate.EditValue.ToString(), dtFormat).ToString("yyyy-MM-dd");
        }


    }
}
