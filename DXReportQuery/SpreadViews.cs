using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data;
using DevExpress.Spreadsheet;

namespace DXReportQuery
{
    partial class SpreadView
    {
        public static void DjwtView()
        {
            DataTable djwtDataTable = QueryResults.DjwtQuery();
            Dictionary<string, int> branchCount= new Dictionary<string, int>();

            var queryCountResult = from djwt in djwtDataTable.AsEnumerable()
                                   group djwt by new { ver = djwt.Field<string>("ver") }
                                   into g
                                   select new
                                   {
                                       g.Key.ver,
                                       count = g.Count()
                                   };

            if(queryCountResult.ToList().Count > 0)
            {
                queryCountResult.ToList().ForEach(q => branchCount.Add(q.ver, q.count));
            }

            IWorkbook workbook = frmMainView.frmMainForm.ssQueryResultView.Document;
            workbook.Worksheets.Add();
        }
    }
}
