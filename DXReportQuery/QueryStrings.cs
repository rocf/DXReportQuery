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

        internal static string ztgblQuery = @"
WITH det
AS ( SELECT t.ver ,
            t.totalNum ,
            t.totalNumNoPay ,
            t.NoStateNumNoPay ,
            t.confirmNum ,
            t.adjusNum ,
            t.waitNum ,
            t.closedNum ,
            t.closeRate AS closeRate ,
            t.payNum
     FROM   (   SELECT   ( CASE dept
                                WHEN '1' THEN '商超'
                                WHEN '2' THEN '餐饮'
                                WHEN '3' THEN '专卖'
                                WHEN '8' THEN '商锐'
                                WHEN '6' THEN 'ESHOP'
                                WHEN 'H' THEN '孕婴童'
                                WHEN 'I' THEN '星食客'
                                WHEN 'C' THEN '新零售'
                                ELSE 'other'
                           END ) AS ver ,
                         COUNT(recno) AS totalNum ,
                         SUM(CASE category
                                  WHEN '6' THEN 0
                                  WHEN '8' THEN 0
                                  ELSE 1
                             END) AS totalNumNoPay ,
                         SUM(CASE category
                                  WHEN '6' THEN 0
                                  WHEN '8' THEN 0
                                  ELSE ( CASE ISNULL(ModifyCode, '1')
                                              WHEN '1' THEN 1
                                              ELSE 0
                                         END )
                             END) AS NoStateNumNoPay ,
                         SUM(CASE category
                                  WHEN '6' THEN 0
                                  WHEN '8' THEN 0
                                  ELSE
                             ( CASE ISNULL(ModifyCode, '1')
                                    WHEN '1' THEN ( CASE Status
                                                         WHEN '1' THEN 1
                                                         ELSE 0
                                                    END )
                                    ELSE 0
                               END )
                             END) confirmNum ,
                         SUM(CASE category
                                  WHEN '6' THEN 0
                                  WHEN '8' THEN 0
                                  ELSE
                             ( CASE ISNULL(ModifyCode, '1')
                                    WHEN '1' THEN ( CASE Status
                                                         WHEN '2' THEN 1
                                                         ELSE 0
                                                    END )
                                    ELSE 0
                               END )
                             END) adjusNum ,
                         SUM(CASE category
                                  WHEN '6' THEN 0
                                  WHEN '8' THEN 0
                                  ELSE
                             ( CASE ISNULL(ModifyCode, '1')
                                    WHEN '1' THEN ( CASE Status
                                                         WHEN '3' THEN 1
                                                         ELSE 0
                                                    END )
                                    ELSE 0
                               END )
                             END) waitNum ,
                         SUM(CASE category
                                  WHEN '6' THEN 0
                                  WHEN '8' THEN 0
                                  ELSE
                             ( CASE ISNULL(ModifyCode, '1')
                                    WHEN '1' THEN ( CASE Status
                                                         WHEN '4' THEN 1
                                                         WHEN '5' THEN 1
                                                         ELSE 0
                                                    END )
                                    ELSE 0
                               END )
                             END) closedNum ,
                         1.0
                         * SUM(CASE category
                                    WHEN '6' THEN 0
                                    WHEN '8' THEN 0
                                    ELSE
                               ( CASE ISNULL(ModifyCode, '1')
                                      WHEN '1' THEN ( CASE Status
                                                           WHEN '4' THEN 1
                                                           WHEN '5' THEN 1
                                                           ELSE 0
                                                      END )
                                      ELSE 0
                                 END )
                               END)
                         / SUM(CASE category
                                    WHEN '6' THEN 0
                                    WHEN '8' THEN 0
                                    ELSE ( CASE ISNULL(ModifyCode, '1')
                                                WHEN '1' THEN 1
                                                ELSE 0
                                           END )
                               END) AS closeRate ,
                         SUM(CASE category
                                  WHEN '6' THEN 1
                                  WHEN '8' THEN 1
                                  ELSE 0
                             END) AS payNum
                FROM     QAQuestion a ,
                         QADeptMaintenance b ,
                         qauser c
                WHERE    a.version = b.version
                         AND b.dept IN ( '1', '2', '3', '6', '8', 'H', 'I', 'C' )
                         AND CONVERT(CHAR(10), a.FirstSubmitDate, 121) >= '{0:G}'
                         AND CONVERT(CHAR(10), a.FirstSubmitDate, 121) <= '{1:G}'
                         AND category <> '2'
                         AND a.userid NOT LIKE 'v%'
                         AND a.userid = c.userid
                --AND isnull(a.IsApproved,'') like :as_class1 百杰 
                --AND isnull(c.class,'') like :as_class 代理商级别 
                GROUP BY dept ) AS t ) ,
     ret
AS ( SELECT   ( CASE dept
                     WHEN '1' THEN '商超'
                     WHEN '2' THEN '餐饮'
                     WHEN '3' THEN '专卖'
                     WHEN '8' THEN '商锐'
                     WHEN '6' THEN 'ESHOP'
                     WHEN 'H' THEN '孕婴童'
                     WHEN 'I' THEN '星食客'
                     WHEN 'C' THEN '新零售'
                     ELSE 'other'
                END ) AS ver ,
              COUNT(recno) AS totalNumLastWeek ,
              1.0 * SUM(CASE category
                             WHEN '6' THEN 0
                             WHEN '8' THEN 0
                             ELSE ( CASE ISNULL(ModifyCode, '1')
                                         WHEN '1' THEN ( CASE Status
                                                              WHEN '4' THEN 1
                                                              WHEN '5' THEN 1
                                                              ELSE 0
                                                         END )
                                         ELSE 0
                                    END )
                        END) / SUM(CASE category
                                        WHEN '6' THEN 0
                                        WHEN '8' THEN 0
                                        ELSE ( CASE ISNULL(ModifyCode, '1')
                                                    WHEN '1' THEN 1
                                                    ELSE 0
                                               END )
                                   END) AS closeRateLastWeek
     FROM     QAQuestion a ,
              QADeptMaintenance b ,
              qauser c
     WHERE    a.version = b.version
              AND b.dept IN ( '1', '2', '3', '6', '8', 'H', 'I', 'C' )
              AND CONVERT(CHAR(10), a.FirstSubmitDate, 121) >= CONVERT(
                                                                   VARCHAR(10) ,
                                                                   DATEADD(
                                                                       DAY ,
                                                                       -7,
                                                                       '{0:G}'),
                                                                   121)
              AND CONVERT(CHAR(10), a.FirstSubmitDate, 121) <= CONVERT(
                                                                   VARCHAR(10) ,
                                                                   DATEADD(
                                                                       DAY ,
                                                                       -7,
                                                                       '{1:G}'),
                                                                   121)
              AND category <> '2'
              AND a.userid NOT LIKE 'v%'
              AND a.userid = c.userid
     --AND isnull(a.IsApproved,'') like :as_class1 百杰 
     --AND isnull(c.class,'') like :as_class 代理商级别 
     GROUP BY dept )
SELECT det.ver ,
       det.totalNum ,
       ret.totalNumLastWeek ,
       det.totalNum - ret.totalNumLastWeek lrr ,
       CONVERT(NUMERIC(16, 4), ( det.totalNum - ret.totalNumLastWeek ))
       / CONVERT(NUMERIC(16, 4), ret.totalNumLastWeek) llrrate ,
       det.totalNumNoPay ,
       det.NoStateNumNoPay ,
       det.totalNum - det.totalNumNoPay totalNumPay ,
       det.confirmNum ,
       det.adjusNum ,
       det.waitNum ,
       det.closedNum ,
       det.closeRate ,
       ret.closeRateLastWeek ,
       det.closeRate - ret.closeRateLastWeek compared
FROM   det
       LEFT JOIN ret ON det.ver = ret.ver
ORDER BY SUBSTRING(det.ver, 1, 1) DESC;
";

        internal static string vipWtgblQuery = @"
WITH det AS
(SELECT t.ver ,
       t.totalNum ,
       t.totalNumNoPay ,
       t.NoStateNumNoPay ,
	   t.payNum,
       t.confirmNum ,
       t.adjusNum ,
       t.waitNum ,
       t.closedNum ,
       t.closeRate       
FROM   (   SELECT   ( CASE dept
                           WHEN '1' THEN '商超'
                           WHEN '2' THEN '餐饮'
                           WHEN '3' THEN '专卖'
                           WHEN '8' THEN '商锐'
                           WHEN '6' THEN 'ESHOP'
                           WHEN 'H' THEN '孕婴童'
                           WHEN 'I' THEN '星食客'
                           WHEN 'C' THEN '新零售'
                           ELSE 'other'
                      END ) AS ver ,
                    COUNT(recno) AS totalNum ,
                    SUM(CASE category
                             WHEN '6' THEN 0
                             WHEN '8' THEN 0
                             ELSE 1
                        END) AS totalNumNoPay ,
                    SUM(CASE category
                             WHEN '6' THEN 0
                             WHEN '8' THEN 0
                             ELSE ( CASE ISNULL(ModifyCode, '1')
                                         WHEN '1' THEN 1
                                         ELSE 0
                                    END )
                        END) AS NoStateNumNoPay ,
                    SUM(CASE category
                             WHEN '6' THEN 0
                             WHEN '8' THEN 0
                             ELSE ( CASE ISNULL(ModifyCode, '1')
                                         WHEN '1' THEN ( CASE Status
                                                              WHEN '1' THEN 1
                                                              ELSE 0
                                                         END )
                                         ELSE 0
                                    END )
                        END) confirmNum ,
                    SUM(CASE category
                             WHEN '6' THEN 0
                             WHEN '8' THEN 0
                             ELSE ( CASE ISNULL(ModifyCode, '1')
                                         WHEN '1' THEN ( CASE Status
                                                              WHEN '2' THEN 1
                                                              ELSE 0
                                                         END )
                                         ELSE 0
                                    END )
                        END) adjusNum ,
                    SUM(CASE category
                             WHEN '6' THEN 0
                             WHEN '8' THEN 0
                             ELSE ( CASE ISNULL(ModifyCode, '1')
                                         WHEN '1' THEN ( CASE Status
                                                              WHEN '3' THEN 1
                                                              ELSE 0
                                                         END )
                                         ELSE 0
                                    END )
                        END) waitNum ,
                    SUM(CASE category
                             WHEN '6' THEN 0
                             WHEN '8' THEN 0
                             ELSE ( CASE ISNULL(ModifyCode, '1')
                                         WHEN '1' THEN ( CASE Status
                                                              WHEN '4' THEN 1
                                                              WHEN '5' THEN 1
                                                              ELSE 0
                                                         END )
                                         ELSE 0
                                    END )
                        END) closedNum ,
                    1.0
                    * SUM(CASE category
                               WHEN '6' THEN 0
                               WHEN '8' THEN 0
                               ELSE ( CASE ISNULL(ModifyCode, '1')
                                           WHEN '1' THEN ( CASE Status
                                                                WHEN '4' THEN 1
                                                                WHEN '5' THEN 1
                                                                ELSE 0
                                                           END )
                                           ELSE 0
                                      END )
                          END) / SUM(CASE category
                                          WHEN '6' THEN 0
                                          WHEN '8' THEN 0
                                          ELSE ( CASE ISNULL(ModifyCode, '1')
                                                      WHEN '1' THEN 1
                                                      ELSE 0
                                                 END )
                                     END) AS closeRate ,
                    SUM(CASE category
                             WHEN '6' THEN 1
                             WHEN '8' THEN 1
                             ELSE 0
                        END) AS payNum
           FROM     QAQuestion a ,
                    QADeptMaintenance b ,
                    qauser c
           WHERE    a.version = b.version
                    AND b.dept IN ( '1', '2', '3', '6', '8', 'H', 'I', 'C' )
                    AND CONVERT(CHAR(10), a.FirstSubmitDate, 121) >= '{0:G}'
                    AND CONVERT(CHAR(10), a.FirstSubmitDate, 121) <= '{1:G}'
                    AND category <> '2'
                    AND a.userid NOT LIKE 'v%'
                    AND a.userid = c.userid
                    AND a.userid IN (   SELECT agentid
                                        FROM   t_AgentManger
                                        WHERE  t_AgentManger.Trade = b.dept
                                               AND Type = '1' )
           GROUP BY dept ) AS t
), 
ret AS 
(
SELECT t.ver ,
       t.closeRate
       
FROM   (   SELECT   ( CASE dept
                           WHEN '1' THEN '商超'
                           WHEN '2' THEN '餐饮'
                           WHEN '3' THEN '专卖'
                           WHEN '8' THEN '商锐'
                           WHEN '6' THEN 'ESHOP'
                           WHEN 'H' THEN '孕婴童'
                           WHEN 'I' THEN '星食客'
                           WHEN 'C' THEN '新零售'
                           ELSE 'other'
                      END ) AS ver ,
                    1.0
                    * SUM(CASE category
                               WHEN '6' THEN 0
                               WHEN '8' THEN 0
                               ELSE ( CASE ISNULL(ModifyCode, '1')
                                           WHEN '1' THEN ( CASE Status
                                                                WHEN '4' THEN 1
                                                                WHEN '5' THEN 1
                                                                ELSE 0
                                                           END )
                                           ELSE 0
                                      END )
                          END) / SUM(CASE category
                                          WHEN '6' THEN 0
                                          WHEN '8' THEN 0
                                          ELSE ( CASE ISNULL(ModifyCode, '1')
                                                      WHEN '1' THEN 1
                                                      ELSE 0
                                                 END )
                                     END) AS closeRate 
           FROM     QAQuestion a ,
                    QADeptMaintenance b ,
                    qauser c
           WHERE    a.version = b.version
                    AND b.dept IN ( '1', '2', '3', '6', '8', 'H', 'I', 'C' )
                    AND CONVERT(CHAR(10), a.FirstSubmitDate, 121) >= CONVERT(
                                                                   VARCHAR(10) ,
                                                                   DATEADD(
                                                                       DAY ,
                                                                       -7,
                                                                       '{0:G}'),
                                                                   121)
                    AND CONVERT(CHAR(10), a.FirstSubmitDate, 121) <= CONVERT(
                                                                   VARCHAR(10) ,
                                                                   DATEADD(
                                                                       DAY ,
                                                                       -7,
                                                                       '{1:G}'),
                                                                   121)
                    AND category <> '2'
                    AND a.userid NOT LIKE 'v%'
                    AND a.userid = c.userid
                    AND a.userid IN (   SELECT agentid
                                        FROM   t_AgentManger
                                        WHERE  t_AgentManger.Trade = b.dept )
           GROUP BY dept ) AS t
),
yet AS 
(
SELECT  t.ver ,
        t.totalNum
FROM    ( SELECT    ( CASE dept
                        WHEN '1' THEN '商超'
                        WHEN '2' THEN '餐饮'
                        WHEN '3' THEN '专卖'
                        WHEN '8' THEN '商锐'
                        WHEN '6' THEN 'ESHOP'
                        WHEN 'H' THEN '孕婴童'
                        WHEN 'I' THEN '星食客'
                        WHEN 'C' THEN '新零售'
                        ELSE 'other'
                      END ) AS ver ,
                    COUNT(recno) AS totalNum 
          FROM      QAQuestion a ,
                    QADeptMaintenance b ,
                    qauser c
          WHERE     a.version = b.version
                    AND b.dept IN ( '1', '2', '3', '6', '8', 'H', 'I', 'C' )
                    AND CONVERT(CHAR(10), a.FirstSubmitDate, 121) >= '{0:G}'
                    AND CONVERT(CHAR(10), a.FirstSubmitDate, 121) <= '{1:G}'
                    AND category <> '2'
                    AND a.userid NOT LIKE 'v%'
                    AND a.userid = c.userid
--AND isnull(a.IsApproved,'') like :as_class1 百杰 
--AND isnull(c.class,'') like :as_class 代理商级别 
GROUP BY            dept
        ) AS t
)
SELECT det.ver, det.totalNum, det.totalNumNoPay, det.NoStateNumNoPay, det.payNum, det.confirmNum, det.adjusNum, det.waitNum,
		det.closedNum, det.closeRate, ret.closeRate, det.closeRate - ret.closeRate comparedRate, (1.0*det.totalNum)/(1.0*yet.totalNum) totalNumRate
FROM det 
left JOIN ret ON ret.ver = det.ver
left JOIN yet ON yet.ver = det.ver AND yet.ver = ret.ver
ORDER BY SUBSTRING(det.ver, 1, 1) 

";

        internal static string clzWtclQuery = @"
WITH det as
(SELECT t.ver ,
        t.totalNum ,
        t.totalNumNoPay ,
        t.NoStateNumNoPay ,
		t.payNum,
        t.confirmNum ,
        t.adjusNum ,
        t.waitNum ,
        t.closedNum ,
        t.closeRate
FROM    ( SELECT    ( CASE dept
                        WHEN '1' THEN '商超'
                        WHEN '2' THEN '餐饮'
                        WHEN '3' THEN '专卖'
                        WHEN '8' THEN '商锐'
                        WHEN '6' THEN 'ESHOP'
                        WHEN 'H' THEN '孕婴童'
                        WHEN 'I' THEN '星食客'
                        WHEN 'C' THEN '新零售'
                        ELSE 'other'
                      END ) AS ver ,
                    COUNT(recno) AS totalNum ,
                    SUM(CASE category
                          WHEN '6' THEN 0
                          WHEN '8' THEN 0
                          ELSE 1
                        END) AS totalNumNoPay ,
                    SUM(CASE category
                          WHEN '6' THEN 0
                          WHEN '8' THEN 0
                          ELSE ( CASE ISNULL(ModifyCode, '1')
                                   WHEN '1' THEN 1
                                   ELSE 0
                                 END )
                        END) AS NoStateNumNoPay ,
                    SUM(CASE category
                          WHEN '6' THEN 0
                          WHEN '8' THEN 0
                          ELSE ( CASE ISNULL(ModifyCode, '1')
                                   WHEN '1' THEN ( CASE Status
                                                     WHEN '1' THEN 1
                                                     ELSE 0
                                                   END )
                                   ELSE 0
                                 END )
                        END) AS confirmNum ,
                    SUM(CASE category
                          WHEN '6' THEN 0
                          WHEN '8' THEN 0
                          ELSE ( CASE ISNULL(ModifyCode, '1')
                                   WHEN '1' THEN ( CASE Status
                                                     WHEN '2' THEN 1
                                                     ELSE 0
                                                   END )
                                   ELSE 0
                                 END )
                        END) AS adjusNum ,
                    SUM(CASE category
                          WHEN '6' THEN 0
                          WHEN '8' THEN 0
                          ELSE ( CASE ISNULL(ModifyCode, '1')
                                   WHEN '1' THEN ( CASE Status
                                                     WHEN '3' THEN 1
                                                     ELSE 0
                                                   END )
                                   ELSE 0
                                 END )
                        END) AS waitNum ,
                    SUM(CASE category
                          WHEN '6' THEN 0
                          WHEN '8' THEN 0
                          ELSE ( CASE ISNULL(ModifyCode, '1')
                                   WHEN '1' THEN ( CASE Status
                                                     WHEN '4' THEN 1
                                                     WHEN '5' THEN 1
                                                     ELSE 0
                                                   END )
                                   ELSE 0
                                 END )
                        END) AS closedNum ,
                    1.0 * SUM(CASE category
                                WHEN '6' THEN 0
                                WHEN '8' THEN 0
                                ELSE ( CASE ISNULL(ModifyCode, '1')
                                         WHEN '1' THEN ( CASE Status
                                                           WHEN '4' THEN 1
                                                           WHEN '5' THEN 1
                                                           ELSE 0
                                                         END )
                                         ELSE 0
                                       END )
                              END) / SUM(CASE category
                                           WHEN '6' THEN 0
                                           WHEN '8' THEN 0
                                           ELSE ( CASE ISNULL(ModifyCode, '1')
                                                    WHEN '1' THEN 1
                                                    ELSE 0
                                                  END )
                                         END) AS closeRate ,
                    SUM(CASE category
                          WHEN '6' THEN 1
                          WHEN '8' THEN 1
                          ELSE 0
                        END) AS payNum
          FROM      QAQuestion a ,
                    QADeptMaintenance b ,
                    qauser c
          WHERE     a.version = b.version
                    AND b.dept IN ( '1', '2', '3', '6', '8', 'H', 'I', 'C' )
                    AND CONVERT(CHAR(10), a.StartDate, 121) >= '{0:G}'
                    AND CONVERT(CHAR(10), a.StartDate, 121) <= '{1:G}'
                    AND category <> '2'
                    AND a.userid NOT LIKE 'v%'
                    AND a.userid = c.userid
--and isnull(a.IsApproved,'') like :as_class1
--and isnull(c.class,'') like :as_class
GROUP BY            dept
        ) AS t 
),
ret AS
(
SELECT  t.ver ,
		t.closeRate
FROM    ( SELECT    ( CASE dept
                        WHEN '1' THEN '商超'
                        WHEN '2' THEN '餐饮'
                        WHEN '3' THEN '专卖'
                        WHEN '8' THEN '商锐'
                        WHEN '6' THEN 'ESHOP'
                        WHEN 'H' THEN '孕婴童'
                        WHEN 'I' THEN '星食客'
                        WHEN 'C' THEN '新零售'
                        ELSE 'other'
                      END ) AS ver ,
                    1.0 * SUM(CASE category
                                WHEN '6' THEN 0
                                WHEN '8' THEN 0
                                ELSE ( CASE ISNULL(ModifyCode, '1')
                                         WHEN '1' THEN ( CASE Status
                                                           WHEN '4' THEN 1
                                                           WHEN '5' THEN 1
                                                           ELSE 0
                                                         END )
                                         ELSE 0
                                       END )
                              END) / SUM(CASE category
                                           WHEN '6' THEN 0
                                           WHEN '8' THEN 0
                                           ELSE ( CASE ISNULL(ModifyCode, '1')
                                                    WHEN '1' THEN 1
                                                    ELSE 0
                                                  END )
                                         END) AS closeRate 
          FROM      QAQuestion a ,
                    QADeptMaintenance b ,
                    qauser c
          WHERE     a.version = b.version
                    AND b.dept IN ( '1', '2', '3', '6', '8', 'H', 'I', 'C' )
                    AND CONVERT(CHAR(10), a.StartDate, 121) >= CONVERT(
                                                                        VARCHAR(10) ,
                                                                        DATEADD(
                                                                            DAY ,
                                                                            -7,
                                                                            '{0:G}'), 121)
                    AND CONVERT(CHAR(10), a.StartDate, 121) <= CONVERT(
                                                                        VARCHAR(10) ,
                                                                        DATEADD(
                                                                            DAY ,
                                                                            -7,
                                                                            '{1:G}'), 121)
                    AND category <> '2'
                    AND a.userid NOT LIKE 'v%'
                    AND a.userid = c.userid
--and isnull(a.IsApproved,'') like :as_class1
--and isnull(c.class,'') like :as_class
GROUP BY            dept
        ) AS t 
)
SELECT det.ver, det.totalNum, det.totalNumNoPay, det.NoStateNumNoPay, det.payNum, det.confirmNum, 
	det.adjusNum, det.waitNum, det.closedNum, det.closeRate,ret.closeRate, det.closeRate - ret.closeRate comparedRate
FROM det 
LEFT JOIN ret ON ret.ver = det.ver
";

        internal static string qyxnQuery = @"
SELECT  ( CASE t4.dept
                WHEN '1' THEN '商超'
                WHEN '2' THEN '餐饮'
                WHEN '3' THEN '专卖'
                WHEN '5' THEN '流通'
				WHEN '6' THEN 'eshop'
                WHEN '8' THEN '商锐'
                WHEN 'H' THEN '孕婴童'
                WHEN 'I' THEN '星食客'
                WHEN 'C' THEN '生鲜便利'
                ELSE '其他'
           END ) AS ver , 
         t3.Name ,
         t1.industry AS provice ,
         COUNT(DISTINCT t2.RecNo) AS totalnum ,
         CAST(ROUND(
                  SUM(DATEDIFF(
                          MINUTE ,
                          t2.FirstSubmitDate,
                          ISNULL(t2.firsthandledate, GETDATE())))
                  / ( COUNT(DISTINCT t2.RecNo) * 60.0 ) ,
                  2) AS NUMERIC(20, 2)) AS real_first ,
         CAST(ROUND(
                  SUM(DATEDIFF(
                          MINUTE ,
                          t2.FirstSubmitDate,
                          ISNULL(t2.finishhandledate, GETDATE())))
                  / ( COUNT(DISTINCT t2.RecNo) * 60.0 ) ,
                  2) AS NUMERIC(20, 2)) AS real_handle ,
         CAST(ROUND(
                  SUM(DATEDIFF(
                          MINUTE ,
                          t2.FirstSubmitDate,
                          ISNULL(t2.Finishdate, GETDATE())))
                  / ( COUNT(DISTINCT t2.RecNo) * 60.0 ) ,
                  2) AS NUMERIC(20, 2)) AS real_close ,
         CAST(ROUND(
                  SUM(DATEDIFF(
                          MINUTE ,
                          t2.ValidFirstSubmitDate,
                          CASE WHEN ISNULL(t2.firsthandledate, GETDATE()) > t2.ValidFirstSubmitDate THEN
                                   ISNULL(t2.firsthandledate, GETDATE())
                               ELSE t2.ValidFirstSubmitDate
                          END)) / ( COUNT(DISTINCT t2.RecNo) * 60.0 ) ,
                  2) AS NUMERIC(20, 2)) AS valid_first ,
         CAST(ROUND(
                  SUM(DATEDIFF(
                          MINUTE ,
                          t2.ValidFirstSubmitDate,
                          CASE WHEN ISNULL(t2.finishhandledate, GETDATE()) > t2.ValidFirstSubmitDate THEN
                                   ISNULL(t2.finishhandledate, GETDATE())
                               ELSE t2.ValidFirstSubmitDate
                          END)) / ( COUNT(DISTINCT t2.RecNo) * 60.0 ) ,
                  2) AS NUMERIC(20, 2)) AS valid_handle ,
         CAST(ROUND(
                  SUM(DATEDIFF(
                          MINUTE ,
                          t2.ValidFirstSubmitDate,
                          CASE WHEN ISNULL(t2.Finishdate, GETDATE()) > t2.ValidFirstSubmitDate THEN
                                   ISNULL(t2.Finishdate, GETDATE())
                               ELSE t2.ValidFirstSubmitDate
                          END)) / ( COUNT(DISTINCT t2.RecNo) * 60.0 ) ,
                  2) AS NUMERIC(20, 2)) AS valid_close ,
         CAST(ROUND(( SUM(Handlenum) * 1.0 / COUNT(DISTINCT t2.RecNo)), 2) AS NUMERIC(20, 2)) AS avghandle ,
         SUM(CASE HistoryTimeout
                  WHEN '1' THEN 1
                  ELSE 0
             END) AS needfirsterror ,
         SUM(CASE TimeoutNum
                  WHEN '1' THEN 1
                  ELSE 0
             END) AS needhandleerror
FROM     iss.QAQuestion t1 ,
(   SELECT *
    FROM   QAQuestionEffect
    WHERE  RecNo NOT IN (   SELECT RecNo
                            FROM   dbo.QAQuestionEffect
                            WHERE  firsthandledate IS NULL
                                   AND Status = '4'
                                   AND CONVERT(CHAR(10), FirstSubmitDate, 121) >= '{0}'
                                   AND CONVERT(CHAR(10), FirstSubmitDate, 121) <= '{1}' )
           AND CONVERT(CHAR(10), FirstSubmitDate, 121) >= '{0}'
           AND CONVERT(CHAR(10), FirstSubmitDate, 121) <= '{1}' ) t2 ,
         iss.QAUser t3 ,
         QADeptMaintenance t4 ,
         QATaskDistribution t5
WHERE    t1.RecNo = t2.RecNo
         AND SUBSTRING(t1.UserID, 1, 1) <> 'v'
         AND CONVERT(CHAR(10), t1.FirstSubmitDate, 121) >= '{0}'
         AND CONVERT(CHAR(10), t1.FirstSubmitDate, 121) <= '{1}'
         AND t1.UserID NOT LIKE 'siss%'
         AND t1.UserID NOT LIKE '9876%'
         AND ISNULL(ModifyCode, 1) LIKE '1'
         AND t5.UserID = t3.UserID
         AND Category <> '2'
         AND t1.Version = t4.Version
         AND t4.dept LIKE '{2}'
         AND t1.industry = t5.provice
         AND t5.deptid LIKE '{2}'
         AND t1.UserID NOT IN (   SELECT agentid
                                  FROM   t_AgentManger
                                  WHERE  t_AgentManger.Trade = t4.dept )
GROUP BY t3.Name ,
         t1.industry ,
         t4.dept
ORDER BY t3.Name
";


    }
}
