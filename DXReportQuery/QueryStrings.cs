using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace DXReportQuery
{
    class QueryStrings
    {
        internal static string djwtQuery = @"
SELECT t.*,
       (CASE
            WHEN t.registerNum = 0 THEN
                '0.00%'
            ELSE
                RTRIM(CONVERT(DECIMAL(18, 2), t.returnNum * 100.0 / t.registerNum)) + '%'
        END
       ) AS returnRate
FROM
(
    SELECT (CASE SUBSTRING(version, 1, 1)
                WHEN '1' THEN
                    '商超'
                WHEN '2' THEN
                    '餐饮'
                WHEN '3' THEN
                    '专卖'
                WHEN '6' THEN
                    'eshop'
            END
           ) AS ver,
           b.userid,
           b.name,
           SUM(   CASE category
                      WHEN '2' THEN
                          0
                      ELSE
                          1
                  END
              ) AS registerNum,
           SUM(   CASE category
                      WHEN '2' THEN
                          0
                      ELSE
                          CASE callbackstatus
                              WHEN '1' THEN
                                  1
                              ELSE
                                  0
                          END
                  END
              ) AS returnNum,
           SUM(   CASE category
                      WHEN '2' THEN
                          1
                      ELSE
                          0
                  END
              ) AS SuggestNum
    FROM qaquestion a,
         qauser b
    WHERE CONVERT(CHAR(10), a.firstsubmitdate, 121) >= '{0:G}'
          AND CONVERT(CHAR(10), a.firstsubmitdate, 121) <= '{1:G}'
          AND LEN(a.addedby) = 4
          AND a.addedby = b.userid
    GROUP BY SUBSTRING(version, 1, 1),
             b.name,
             b.userid
) AS t
WHERE ver IS NOT NULL
      AND userid <> 'siss'
ORDER BY SUBSTRING(t.ver, 1, 1) DESC;
";

        internal static string wtgblZtwtgblQuery = @"
SELECT t.ver,
       t.totalNum,
       t.totalNumNoPay,
       t.NoStateNumNoPay,
       t.confirmNum,
       t.adjusNum,
       t.waitNum,
       t.closedNum,
       RTRIM(CONVERT(DECIMAL(18, 2), t.closeRate * 100.0)) + '%' AS closeRate,
       t.payNum
FROM
(
    SELECT (CASE dept
                WHEN '1' THEN
                    '商超'
                WHEN '2' THEN
                    '餐饮'
                WHEN '3' THEN
                    '专卖'
                WHEN '8' THEN
                    '商锐'
                WHEN '6' THEN
                    'ESHOP'
                WHEN 'H' THEN
                    '孕婴童'
                WHEN 'I' THEN
                    '星食客'
                WHEN 'C' THEN
                    '新零售'
                ELSE
                    'other'
            END
           ) AS ver,
           COUNT(recno) AS totalNum,
           SUM(   CASE category
                      WHEN '6' THEN
                          0
                      WHEN '8' THEN
                          0
                      ELSE
                          1
                  END
              ) AS totalNumNoPay,
           SUM(   CASE category
                      WHEN '6' THEN
                          0
                      WHEN '8' THEN
                          0
                      ELSE
                  (CASE ISNULL(ModifyCode, '1')
                       WHEN '1' THEN
                           1
                       ELSE
                           0
                   END
                  )
                  END
              ) AS NoStateNumNoPay,
           SUM(   CASE category
                      WHEN '6' THEN
                          0
                      WHEN '8' THEN
                          0
                      ELSE
                  (CASE ISNULL(ModifyCode, '1')
                       WHEN '1' THEN
                  (CASE Status
                       WHEN '1' THEN
                           1
                       ELSE
                           0
                   END
                  )
                       ELSE
                           0
                   END
                  )
                  END
              ) confirmNum,
           SUM(   CASE category
                      WHEN '6' THEN
                          0
                      WHEN '8' THEN
                          0
                      ELSE
                  (CASE ISNULL(ModifyCode, '1')
                       WHEN '1' THEN
                  (CASE Status
                       WHEN '2' THEN
                           1
                       ELSE
                           0
                   END
                  )
                       ELSE
                           0
                   END
                  )
                  END
              ) adjusNum,
           SUM(   CASE category
                      WHEN '6' THEN
                          0
                      WHEN '8' THEN
                          0
                      ELSE
                  (CASE ISNULL(ModifyCode, '1')
                       WHEN '1' THEN
                  (CASE Status
                       WHEN '3' THEN
                           1
                       ELSE
                           0
                   END
                  )
                       ELSE
                           0
                   END
                  )
                  END
              ) waitNum,
           SUM(   CASE category
                      WHEN '6' THEN
                          0
                      WHEN '8' THEN
                          0
                      ELSE
                  (CASE ISNULL(ModifyCode, '1')
                       WHEN '1' THEN
                  (CASE Status
                       WHEN '4' THEN
                           1
                       WHEN '5' THEN
                           1
                       ELSE
                           0
                   END
                  )
                       ELSE
                           0
                   END
                  )
                  END
              ) closedNum,
           1.0 * SUM(   CASE category
                            WHEN '6' THEN
                                0
                            WHEN '8' THEN
                                0
                            ELSE
                        (CASE ISNULL(ModifyCode, '1')
                             WHEN '1' THEN
                        (CASE Status
                             WHEN '4' THEN
                                 1
                             WHEN '5' THEN
                                 1
                             ELSE
                                 0
                         END
                        )
                             ELSE
                                 0
                         END
                        )
                        END
                    ) / SUM(   CASE category
                                   WHEN '6' THEN
                                       0
                                   WHEN '8' THEN
                                       0
                                   ELSE
                               (CASE ISNULL(ModifyCode, '1')
                                    WHEN '1' THEN
                                        1
                                    ELSE
                                        0
                                END
                               )
                               END
                           ) AS closeRate,
           SUM(   CASE category
                      WHEN '6' THEN
                          1
                      WHEN '8' THEN
                          1
                      ELSE
                          0
                  END
              ) AS payNum
    FROM QAQuestion a,
         QADeptMaintenance b,
         qauser c
    WHERE a.version = b.version
          AND b.dept IN ( '1', '2', '3', '6', '8', 'H', 'I', 'C' )
          AND CONVERT(CHAR(10), a.FirstSubmitDate, 121) >= '{0:G}'
          AND CONVERT(CHAR(10), a.FirstSubmitDate, 121) <= '{1:G}'
          AND category <> '2'
          AND a.userid NOT LIKE 'v%'
          AND a.userid = c.userid
    --AND isnull(a.IsApproved,'') like :as_class1 百杰 
    --AND isnull(c.class,'') like :as_class 代理商级别 
    GROUP BY dept
) AS t;
";
    }
}
