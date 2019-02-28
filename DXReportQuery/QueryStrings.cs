using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace DXReportQuery
{
    class QueryStrings
    {
        internal static string djwtQuery = @"
SELECT t.*, (CASE WHEN t.registerNum=0 THEN '0.00%' ELSE RTRIM(CONVERT(DECIMAL(18, 2), t.returnNum * 100.0 / t.registerNum))+'%' END) AS returnRate
FROM(SELECT (CASE SUBSTRING(version, 1, 1)WHEN '1' THEN '商超'
             WHEN '2' THEN '餐饮'
             WHEN '3' THEN '专卖'
             WHEN '6' THEN 'eshop' END) AS ver, b.userid, b.name, SUM(CASE category WHEN '2' THEN 0 ELSE 1 END) AS registerNum, SUM(CASE category WHEN '2' THEN 0 ELSE CASE callbackstatus WHEN '1' THEN 1 ELSE 0 END END) AS returnNum, SUM(CASE category WHEN '2' THEN 1 ELSE 0 END) AS SuggestNum
     FROM qaquestion a, qauser b
     WHERE CONVERT(CHAR(10), a.firstsubmitdate, 121)>='{0:G}' AND CONVERT(CHAR(10), a.firstsubmitdate, 121)<='{1:G}' AND LEN(a.addedby)=4 AND a.addedby=b.userid
     GROUP BY SUBSTRING(version, 1, 1), b.name, b.userid) AS t
WHERE ver IS NOT null
ORDER BY SUBSTRING(t.ver, 1, 1);";


    }
}
