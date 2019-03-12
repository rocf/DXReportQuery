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
                                WHEN 'C' THEN '生鲜便利'
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
                     WHEN 'C' THEN '生鲜便利'
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
                           WHEN 'C' THEN '生鲜便利'
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
                           WHEN 'C' THEN '生鲜便利'
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
                        WHEN 'C' THEN '生鲜便利'
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
                        WHEN 'C' THEN '生鲜便利'
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
                        WHEN 'C' THEN '生鲜便利'
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

        internal static string vipGblQuery = @"
SELECT   t.ver ,
         t.name ,         
         t.companyName ,
         t.regionQuesNum ,
         t.NoModifQues ,
         t.handerNum ,
         ( t.NoModifQues - t.handerNum ) AS assistDealNum ,
         t.closedNum ,
         ( CASE WHEN t.NoModifQues = 0 THEN '0.00%'
                ELSE
                    RTRIM(
                        CONVERT(
                            DECIMAL(18, 2), t.closedNum * 100.0 / t.NoModifQues))
                    + '%'
           END ) AS regionClosedRate ,
         ( CASE WHEN t.NoModifQues = 0 THEN '0.00%'
                ELSE
                    RTRIM(
                        CONVERT(
                            DECIMAL(18, 2), t.handerNum * 100.0 / t.NoModifQues))
                    + '%'
           END ) AS handRate
FROM     (   SELECT   qauser.name ,
                      ( CASE SUBSTRING(QAQuestion.version, 1, 1)
                             WHEN '1' THEN '商超'
                             WHEN '2' THEN '餐饮'
                             WHEN '3' THEN '专卖'
                             WHEN '6' THEN 'ESHOP'
                             WHEN 'H' THEN '孕婴童'
                             WHEN 'I' THEN '星食客'
                             WHEN 'C' THEN '生鲜便利'
                             ELSE ''
                        END ) AS ver ,
                      SUM(CASE ISNULL(ModifyCode, 1)
                               WHEN '1' THEN 1
                               ELSE 0
                          END) AS NoModifQues ,
                      COUNT(*) AS regionQuesNum ,
                      SUM(CASE ISNULL(modifycode, 1)
                               WHEN '1' THEN
                          ( CASE handler
                                 WHEN dealwither THEN ( CASE ISNULL(ModifyCode, 1)
                                                             WHEN '1' THEN 1
                                                             ELSE 0
                                                        END )
                                 ELSE 0
                            END )
                               ELSE 0
                          END) AS handerNum ,
                      SUM(CASE ISNULL(modifycode, 1)
                               WHEN '1' THEN ( CASE Status
                                                    WHEN '4' THEN 1
                                                    WHEN '5' THEN 1
                                                    ELSE 0
                                               END )
                               ELSE 0
                          END) AS closedNum ,
                      b.Name AS companyName
             FROM     QAQuestion ,
                      QADeptMaintenance ,
                      qauser ,
                      qauser AS b
             WHERE    SUBSTRING(QAQuestion.userid, 1, 1) <> 'v'
                      AND addedby NOT LIKE 'siss%'
                      AND QAQuestion.userid NOT LIKE 'siss%'
                      AND QAQuestion.userid NOT LIKE '9876'
                      AND QAQuestion.pointdealwither = qauser.userid
                      AND QAQuestion.userid = b.userid
                      AND CONVERT(CHAR(10), QAQuestion.FirstSubmitDate, 121) >= '{0}'
                      AND CONVERT(CHAR(10), QAQuestion.FirstSubmitDate, 121) <= '{1}'
                      AND (   Category <> '2'
                              AND Category <> '6'
                              AND category <> '8' )
                      AND dealwither IS NOT NULL
                      AND QADeptMaintenance.dept IN ( '1', '2', '3', '6', '8', 'H' ,
                                                      'I' , 'C' )
                      AND QAQuestion.version = QADeptMaintenance.version
                      AND (   (   QADeptMaintenance.dept IN ( '1', '3' )
                                  AND '{2}' = '5' )
                              OR QADeptMaintenance.dept = '{2}' )
                      AND QAQuestion.userid IN (   SELECT agentid
                                                   FROM   t_AgentManger
                                                   WHERE  trade = QADeptMaintenance.dept
                                                          AND Type = '1' ) ---VIP代理商编号
                      AND QAQuestion.pointdealwither IN (   SELECT UserId
                                                            FROM   t_AgentManger
                                                            WHERE  trade = QADeptMaintenance.dept
                                                                   AND Type = '1' ) ------负责人工号                               
                      AND b.Name LIKE '%'
                      AND qauser.UserID LIKE '%'
             GROUP BY QAQuestion.industry ,
                      qauser.name ,
                      b.Name ,
                      QAQuestion.userid ,
                      SUBSTRING(QAQuestion.version, 1, 1)) AS t
ORDER BY t.name;

";

        internal static string qybbQuery = @"
SELECT   t.name ,
         t.ver ,
         t.provice ,
         t.regionQuesNum ,
         t.NoModifQues ,
         t.handerNum ,
         ( t.NoModifQues - t.handerNum ) AS assistDealNum ,
         t.closedNum ,
         ( CASE WHEN t.NoModifQues = 0 THEN '0.00%'
                ELSE
                    RTRIM(
                        CONVERT(
                            DECIMAL(18, 2), t.closedNum * 100.0 / t.NoModifQues))
                    + '%'
           END ) AS regionClosedRate ,
         ( CASE WHEN t.NoModifQues = 0 THEN '0.00%'
                ELSE
                    RTRIM(
                        CONVERT(
                            DECIMAL(18, 2), t.handerNum * 100.0 / t.NoModifQues))
                    + '%'
           END ) AS handRate
FROM     (   SELECT   qauser.name ,
                      ( CASE SUBSTRING(QAQuestion.version, 1, 1)
                             WHEN '1' THEN '商超'
                             WHEN '2' THEN '餐饮'
                             WHEN '3' THEN '专卖'
                             WHEN '6' THEN 'ESHOP'
                             WHEN 'H' THEN '孕婴童'
                             WHEN 'I' THEN '星食客'
                             WHEN 'C' THEN '生鲜便利'
                             ELSE ''
                        END ) AS ver ,
                      QAQuestion.industry AS provice ,
                      COUNT(*) AS regionQuesNum ,
                      SUM(CASE ISNULL(ModifyCode, 1)
                               WHEN '1' THEN 1
                               ELSE 0
                          END) AS NoModifQues ,
                      SUM(CASE ISNULL(modifycode, 1)
                               WHEN '1' THEN
                          ( CASE handler
                                 WHEN dealwither THEN ( CASE ISNULL(ModifyCode, 1)
                                                             WHEN '1' THEN 1
                                                             ELSE 0
                                                        END )
                                 ELSE 0
                            END )
                               ELSE 0
                          END) AS handerNum ,
                      SUM(CASE ISNULL(modifycode, 1)
                               WHEN '1' THEN ( CASE Status
                                                    WHEN '4' THEN 1
                                                    WHEN '5' THEN 1
                                                    ELSE 0
                                               END )
                               ELSE 0
                          END) AS closedNum
             FROM     QAQuestion ,
                      QADeptMaintenance ,
                      qauser
             WHERE    SUBSTRING(QAQuestion.userid, 1, 1) <> 'v'
                      AND addedby NOT LIKE 'siss%'
                      AND QAQuestion.userid NOT LIKE 'siss%'
                      AND QAQuestion.userid NOT LIKE '9876'
                      AND QAQuestion.dealwither = qauser.userid
                      AND CONVERT(CHAR(10), QAQuestion.FirstSubmitDate, 121) >= '{0}'
                      AND CONVERT(CHAR(10), QAQuestion.FirstSubmitDate, 121) <= '{1}'
                      AND (   Category <> '2'
                              AND Category <> '6'
                              AND category <> '8' )
                      AND dealwither IS NOT NULL
                      AND QADeptMaintenance.dept IN ( '1', '2', '3', '6', '8', 'H' ,
                                                      'I' , 'C' )
                      AND QAQuestion.version = QADeptMaintenance.version
                      AND (   (   QADeptMaintenance.dept IN ( '1', '3' )
                                  AND '{2}' = '5' )
                              OR QADeptMaintenance.dept = '{2}' )
                      AND QAQuestion.UserID NOT IN (   SELECT agentid
                                                       FROM   t_AgentManger
                                                       WHERE  t_AgentManger.Trade = QADeptMaintenance.dept )
             GROUP BY QAQuestion.industry ,
                      qauser.name ,
                      SUBSTRING(QAQuestion.version, 1, 1)) AS t
ORDER BY t.name;
";

        internal static string grxnQuery = @"
SELECT   ( CASE t4.dept
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
           END ) AS companyName ,
         t3.name ,
         COUNT(DISTINCT t2.recno) AS totalnum ,
         CAST(ROUND(
                  SUM(DATEDIFF(
                          MINUTE ,
                          t2.firstsubmitdate,
                          ISNULL(t2.firsthandledate, GETDATE())))
                  / ( COUNT(DISTINCT t2.recno) * 60.0 ) ,
                  2) AS NUMERIC(20, 2)) AS real_first ,
         CAST(ROUND(
                  SUM(DATEDIFF(
                          MINUTE ,
                          t2.firstsubmitdate,
                          ISNULL(t2.finishhandledate, GETDATE())))
                  / ( COUNT(DISTINCT t2.recno) * 60.0 ) ,
                  2) AS NUMERIC(20, 2)) AS real_handle ,
         CAST(ROUND(
                  SUM(DATEDIFF(
                          MINUTE ,
                          t2.firstsubmitdate,
                          ISNULL(t2.finishdate, GETDATE())))
                  / ( COUNT(DISTINCT t2.recno) * 60.0 ) ,
                  2) AS NUMERIC(20, 2)) AS real_close ,
         CAST(ROUND(
                  SUM(DATEDIFF(
                          MINUTE ,
                          t2.validfirstsubmitdate,
                          CASE WHEN ISNULL(t2.firsthandledate, GETDATE()) > t2.validfirstsubmitdate THEN
                                   ISNULL(t2.firsthandledate, GETDATE())
                               ELSE t2.validfirstsubmitdate
                          END)) / ( COUNT(DISTINCT t2.recno) * 60.0 ) ,
                  2) AS NUMERIC(20, 2)) AS valid_first ,
         CAST(ROUND(
                  SUM(DATEDIFF(
                          MINUTE ,
                          t2.validfirstsubmitdate,
                          CASE WHEN ISNULL(t2.finishhandledate, GETDATE()) > t2.validfirstsubmitdate THEN
                                   ISNULL(t2.finishhandledate, GETDATE())
                               ELSE t2.validfirstsubmitdate
                          END)) / ( COUNT(DISTINCT t2.recno) * 60.0 ) ,
                  2) AS NUMERIC(20, 2)) AS valid_handle ,
         CAST(ROUND(
                  SUM(DATEDIFF(
                          MINUTE ,
                          t2.validfirstsubmitdate,
                          CASE WHEN ISNULL(t2.finishdate, GETDATE()) > t2.validfirstsubmitdate THEN
                                   ISNULL(t2.finishdate, GETDATE())
                               ELSE t2.validfirstsubmitdate
                          END)) / ( COUNT(DISTINCT t2.recno) * 60.0 ) ,
                  2) AS NUMERIC(20, 2)) AS valid_close ,
         CAST(ROUND(( SUM(handlenum) * 1.0 / COUNT(DISTINCT t2.recno)), 2) AS NUMERIC(20, 2)) AS avghandle ,
         SUM(CASE historytimeout
                  WHEN '1' THEN 1
                  ELSE 0
             END) AS needfirsterror ,
         SUM(CASE timeoutnum
                  WHEN '1' THEN 1
                  ELSE 0
             END) AS needhandleerror
FROM     iss.QAQuestion t1 ,
(   SELECT *
    FROM   QAQuestionEffect
    WHERE  recno NOT IN (   SELECT recno
                            FROM   dbo.QAQuestionEffect
                            WHERE  firsthandledate IS NULL
                                   AND status = '4'
                                   AND CONVERT(CHAR(10), FirstSubmitDate, 121) >= '{0}'
                                   AND CONVERT(CHAR(10), FirstSubmitDate, 121) <= '{1}' )
           AND CONVERT(CHAR(10), FirstSubmitDate, 121) >= '{0}'
           AND CONVERT(CHAR(10), FirstSubmitDate, 121) <= '{1}' ) t2 ,
         iss.QAUser t3 ,
         QADeptMaintenance t4 ,
         QATaskDistribution t5
WHERE    t1.recno = t2.recno
         AND SUBSTRING(t1.userid, 1, 1) <> 'v'
         AND CONVERT(CHAR(10), t1.FirstSubmitDate, 121) >= '{0}'
         AND CONVERT(CHAR(10), t1.FirstSubmitDate, 121) <= '{1}'
         AND DATEPART(WEEKDAY, t1.FirstSubmitDate) <> 1
         AND DATEPART(WEEKDAY, t1.FirstSubmitDate) <> 7
         AND t1.userid NOT LIKE 'siss%'
         AND t1.userid NOT LIKE '9876%'
         AND ISNULL(ModifyCode, 1) LIKE '1'
         AND t5.userid = t3.userid
         AND category <> '2'
         AND t1.version = t4.version
         AND t4.dept LIKE '{2}'
         AND t1.industry = t5.provice
         AND t5.deptid LIKE '{2}'
         AND t1.DealWither = t3.userid
GROUP BY t3.name ,
         t4.dept;
";

        internal static string vipKhjlGrxnQuery = @"
SELECT   ( CASE t4.trade
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
           END ) AS companyName,
		 t3.name ,
         COUNT(DISTINCT t2.recno) AS totalnum ,
         CAST(ROUND(
                  SUM(DATEDIFF(
                          MINUTE ,
                          t2.firstsubmitdate,
                          ISNULL(t2.firsthandledate, GETDATE())))
                  / ( COUNT(DISTINCT t2.recno) * 60.0 ) ,
                  2) AS NUMERIC(20, 2)) AS real_first ,
         CAST(ROUND(
                  SUM(DATEDIFF(
                          MINUTE ,
                          t2.firstsubmitdate,
                          ISNULL(t2.finishhandledate, GETDATE())))
                  / ( COUNT(DISTINCT t2.recno) * 60.0 ) ,
                  2) AS NUMERIC(20, 2)) AS real_handle ,
         CAST(ROUND(
                  SUM(DATEDIFF(
                          MINUTE ,
                          t2.firstsubmitdate,
                          ISNULL(t2.finishdate, GETDATE())))
                  / ( COUNT(DISTINCT t2.recno) * 60.0 ) ,
                  2) AS NUMERIC(20, 2)) AS real_close ,
         CAST(ROUND(
                  SUM(DATEDIFF(
                          MINUTE ,
                          t2.validfirstsubmitdate,
                          CASE WHEN ISNULL(t2.firsthandledate, GETDATE()) > t2.validfirstsubmitdate THEN
                                   ISNULL(t2.firsthandledate, GETDATE())
                               ELSE t2.validfirstsubmitdate
                          END)) / ( COUNT(DISTINCT t2.recno) * 60.0 ) ,
                  2) AS NUMERIC(20, 2)) AS valid_first ,
         CAST(ROUND(
                  SUM(DATEDIFF(
                          MINUTE ,
                          t2.validfirstsubmitdate,
                          CASE WHEN ISNULL(t2.finishhandledate, GETDATE()) > t2.validfirstsubmitdate THEN
                                   ISNULL(t2.finishhandledate, GETDATE())
                               ELSE t2.validfirstsubmitdate
                          END)) / ( COUNT(DISTINCT t2.recno) * 60.0 ) ,
                  2) AS NUMERIC(20, 2)) AS valid_handle ,
         CAST(ROUND(
                  SUM(DATEDIFF(
                          MINUTE ,
                          t2.validfirstsubmitdate,
                          CASE WHEN ISNULL(t2.finishdate, GETDATE()) > t2.validfirstsubmitdate THEN
                                   ISNULL(t2.finishdate, GETDATE())
                               ELSE t2.validfirstsubmitdate
                          END)) / ( COUNT(DISTINCT t2.recno) * 60.0 ) ,
                  2) AS NUMERIC(20, 2)) AS valid_close ,
         CAST(ROUND(( SUM(handlenum) * 1.0 / COUNT(DISTINCT t2.recno)), 2) AS NUMERIC(20, 2)) AS avghandle ,
         SUM(CASE historytimeout
                  WHEN '1' THEN 1
                  ELSE 0
             END) AS needfirsterror ,
         SUM(CASE timeoutnum
                  WHEN '1' THEN 1
                  ELSE 0
             END) AS needhandleerror
FROM     iss.QAQuestion t1 ,
(   SELECT *
    FROM   QAQuestionEffect
    WHERE  recno NOT IN (   SELECT recno
                            FROM   dbo.QAQuestionEffect
                            WHERE  firsthandledate IS NULL
                                   AND status = '4'
                                   AND CONVERT(CHAR(10), FirstSubmitDate, 121) >= '{0}'
                                   AND CONVERT(CHAR(10), FirstSubmitDate, 121) <= '{1}' )
           AND CONVERT(CHAR(10), FirstSubmitDate, 121) >= '{0}'
           AND CONVERT(CHAR(10), FirstSubmitDate, 121) <= '{1}' ) t2 ,
         iss.QAUser t3 ,
         t_AgentManger t4 ,
         iss.QADeptMaintenance t5
WHERE    t1.recno = t2.recno
         AND SUBSTRING(t1.userid, 1, 1) <> 'v'
         AND CONVERT(CHAR(10), t1.FirstSubmitDate, 121) >= '{0}'
         AND CONVERT(CHAR(10), t1.FirstSubmitDate, 121) <= '{1}'
         AND DATEPART(WEEKDAY, t1.FirstSubmitDate) <> 1
         AND DATEPART(WEEKDAY, t1.FirstSubmitDate) <> 7
         AND t1.userid NOT LIKE 'siss%'
         AND t1.userid NOT LIKE '9876%'
         AND ISNULL(ModifyCode, 1) LIKE '1'
         AND t4.userid = t3.userid
         AND category <> '2'
         AND t4.agentid = t1.UserID
         AND t1.version = t5.version
         AND t4.TYPE = '1'
         AND t4.trade LIKE   '{2}'    AND t5.dept LIKE  '{2}' 
GROUP BY t4.trade ,
         t4.USERid ,
         t3.NAME;

";
        internal static string dlsyjQuery = @"
select 
( CASE t5.dept
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
           END ) AS companyName ,
         t2.name AS agent ,
         t3.name AS versionname ,
         (   SELECT name
             FROM   qauser
             WHERE  QAUser.UserID = t4.userid ) AS username ,
         COUNT(*) AS total_num
FROM     iss.QAQuestion t1 ,
         qauser t2 ,
         QADictionary t3 ,
         QATaskDistribution t4 ,
         QADeptMaintenance t5
WHERE    t1.FirstSubmitDate >= '{0}'
         AND t1.FirstSubmitDate <= '{1}'
         AND t1.userid = t2.userid
         AND t2.name NOT LIKE 'siss%'
         AND t3.type = '2'
         AND t3.no = t1.version
         AND t5.dept LIKE '{2}'
         AND t5.version = t1.version
         AND t4.deptid LIKE '{2}'
         AND t2.province = t4.provice
         AND t1.UserID NOT IN (   SELECT agentid
                                  FROM   t_AgentManger
                                  WHERE  t_AgentManger.Trade = t5.dept )
GROUP BY t2.name ,
         t3.name ,
         t4.UserID ,
         t5.dept
HAVING   COUNT(*) >= 5;
";

        internal static string vipDlsyjQuery = @"
SELECT  ( CASE QADeptMaintenance.dept
                WHEN '1' THEN '商超'
                WHEN '2' THEN '餐饮'
                WHEN '3' THEN '专卖'
                WHEN '6' THEN 'ESHOP'
                WHEN 'H' THEN '孕婴童'
                WHEN 'I' THEN '星食客'
                WHEN 'C' THEN '生鲜便利'
                ELSE ''
           END ) AS dept, 
        b.Name AS companyName ,
         ( CASE SUBSTRING(QAQuestion.Version, 1, 1)
                WHEN '1' THEN '商超'
                WHEN '2' THEN '餐饮'
                WHEN '3' THEN '专卖'
                WHEN '6' THEN 'ESHOP'
                WHEN 'H' THEN '孕婴童'
                WHEN 'I' THEN '星食客'
                WHEN 'C' THEN '生鲜便利'
                ELSE ''
           END ) AS ver ,
         QAUser.Name ,
         COUNT(*) AS regionQuesNum
FROM     QAQuestion ,
         QADeptMaintenance ,
         qauser ,
         qauser AS b
WHERE    SUBSTRING(QAQuestion.UserID, 1, 1) <> 'v'
         AND addedby NOT LIKE 'siss%'
         AND QAQuestion.UserID NOT LIKE 'siss%'
         AND QAQuestion.UserID NOT LIKE '9876'
         AND QAQuestion.PointDealWither = QAUser.UserID
         AND QAQuestion.UserID = b.userid
         AND CONVERT(CHAR(10), QAQuestion.FirstSubmitDate, 121) >= '{0}'
         AND CONVERT(CHAR(10), QAQuestion.FirstSubmitDate, 121) <= '{1}'
         AND Category <> '2'
         AND dealwither IS NOT NULL
         AND QADeptMaintenance.dept IN ( '1', '2', '3', '6', '8', 'H', 'I', 'C' )
         AND QAQuestion.Version = QADeptMaintenance.version
         AND (   (   QADeptMaintenance.dept IN ( '1', '3' )
                     AND '{2}' = '5' )
                 OR QADeptMaintenance.dept = '{2}' )
         AND QAQuestion.UserID IN (   SELECT agentid
                                      FROM   t_AgentManger
                                      WHERE  trade = QADeptMaintenance.dept
                                             AND Type = '1' )
         --VIP代理商编号                             
         AND QAQuestion.PointDealWither IN (   SELECT UserId
                                               FROM   t_AgentManger
                                               WHERE  trade = QADeptMaintenance.dept
                                                      AND Type = '1' )
------负责人工号                     
GROUP BY QAQuestion.industry ,
         QAUser.Name ,
         b.Name ,
         QAQuestion.UserID ,
         QADeptMaintenance.dept,
         SUBSTRING(QAQuestion.Version, 1, 1)
HAVING   COUNT(*) >= 5;

";

        internal static string wtyjQuery = @"
SELECT   CASE t4.dept
              WHEN '1' THEN '商云'
              WHEN '2' THEN '餐饮'
              WHEN '3' THEN '专卖'
              WHEN '6' THEN 'ESHOP'
              WHEN 'I' THEN '星食客'
              WHEN 'H' THEN '孕婴童'
              WHEN 'C' THEN '生鲜便利'
              WHEN '8' THEN '商锐'
              ELSE '其他'
         END AS version ,
         (   SELECT name
             FROM   iss.QADictionary
             WHERE  type = 2
                    AND no = t1.version ) AS vername ,
         (   SELECT TOP 1 name
             FROM   qauser
             WHERE  dealwither = userid ) AS username ,
         t1.RecNo ,
         t3.province ,
         CASE ISNULL(t1.Status, 1)
              WHEN '1' THEN '待用户确认'
              WHEN '2' THEN '处理中'
              WHEN '3' THEN '待处理'
              WHEN '4' THEN '关闭'
              WHEN '5' THEN '已评价'
              ELSE '其他'
         END AS Qastatus ,
        /*
		 CASE ISNULL(ModifyCode, 1)
              WHEN '1' THEN '不修改'
              ELSE '需修改'
         END AS modifycode ,
         */
		 CASE develop
              WHEN '1' THEN '是'
              ELSE '否'
         END AS develop ,
         t2.FirstSubmitDate ,
         /*
		 CASE t1.ToDevelopDate
              WHEN NULL THEN '无'
              ELSE t1.ToDevelopDate
         END AS toDevelopDate ,
		 */
         CONVERT(
             DECIMAL(18, 2) ,
             DATEDIFF(
                 MINUTE ,
                 t1.ToDevelopDate,
                 DATENAME(YEAR, GETDATE()) + '-' + DATENAME(MONTH, GETDATE())
                 + '-' + DATENAME(DAY, GETDATE())) / 60.0) AS todevelopTime ,
         t2.finishhandledate ,
         CONVERT(
             DECIMAL(18, 2) ,
             DATEDIFF(MINUTE, t2.FirstSubmitDate, t2.finishhandledate) / 60.0) AS solveTime
FROM     QAQuestionEffect t2 ,
         QAQuestion t1 ,
         qauser t3 ,
         QADeptMaintenance t4
WHERE    t1.RecNo = t2.RecNo
         AND CONVERT(CHAR(10), startdate, 121) >= '{0}'
         AND CONVERT(CHAR(10), startdate, 121) <= '{1}'
         AND t2.firsthandledate IS NOT NULL
         AND DATEDIFF(MINUTE, t2.FirstSubmitDate, t2.finishhandledate) > 24 * 60
                                                                         * 14
         AND t1.userid = t3.userid
         AND ISNULL(t1.modifycode, '1') = '1'
         AND t2.Status <> '4'
         AND category <> '2'
         AND t4.dept IN ( '1', '2', '3', '6', '8', 'H', 'I', 'C' )
         AND t1.version = t4.version
ORDER BY version DESC ,
         vername DESC ,
         username;
";

        internal static string wtxqzblQuery = @"
SELECT  ---t1.Version AS version ,
        CASE t2.dept
          WHEN '1' THEN '商云'
          WHEN '2' THEN '餐饮'
          WHEN '3' THEN '专卖'
          WHEN '6' THEN 'ESHOP'
          WHEN 'I' THEN '星食客'
          WHEN 'H' THEN '孕婴童'
          WHEN 'C' THEN '生鲜便利'
          WHEN '8' THEN '商锐'
          ELSE '其他'
        END AS Dept ,
        ( SELECT    name
          FROM      iss.QADictionary
          WHERE     type = 2
                    AND no = t1.version
        ) AS version ,
        t1.FirstSubmitDate AS FirstSubmitDate ,
        t1.Category AS Category ,
        t1.SubStatus AS ModifyCode ,
        t1.RequirementNo AS RequirementNo ,
        t2.dept AS Dept  
FROM    QAQuestion t1 ,
        QADeptMaintenance t2 
WHERE 
         SUBSTRING(t1.userid, 1, 1) <> 'v'
        AND CONVERT(CHAR(10), t1.FirstSubmitDate, 121) >=  '{0}'
        AND CONVERT(CHAR(10), t1.FirstSubmitDate, 121) <=  '{1}'
        AND t2.dept IN ( '1', '2', '3', '6', '8', 'H', 'I', 'C' )
        AND t1.Version = t2.Version
        ----  AND t2.dept = 'I'
ORDER BY  t2.Dept,t1.Version 
";

        internal static string zzsktjQuery = @"
SELECT   ( CASE c.dept
                WHEN '1' THEN '商超'
                WHEN '2' THEN '餐饮'
                WHEN '3' THEN '专卖'
                WHEN '5' THEN '流通'
                WHEN '6' THEN 'ESHOP'
                WHEN '8' THEN '商锐'
                WHEN 'I' THEN '星食客'
                WHEN 'C' THEN '生鲜便利'
                WHEN 'H' THEN '孕婴童'
                ELSE '其他'
           END ) AS ver ,
         b.name AS name ,
         COUNT(*) AS TransNum
FROM     QAToSissKC a ,
         QAUser b ,
(   SELECT dept ,
           version
    FROM   QADeptMaintenance
    WHERE  dept IN ( '1', '2', '3', '6', '8', 'I', 'C', 'H' )) c ,
         QAQuestion d
WHERE    a.turn_user = b.userid
         AND a.recno = d.recno
         AND d.version = c.version
         AND CONVERT(CHAR(10), turn_date, 121) >= '{0}'
         AND CONVERT(CHAR(10), turn_date, 121) <= '{1}'
GROUP BY b.userid ,
         b.name ,
         c.dept
ORDER BY SUBSTRING(dept, 1, 1),
		 b.name ,
         b.userid ,
         c.dept		 
";

        internal static string zskclsltjQuery = @"
SELECT   name ,
         SUM(CASE ISNULL(faq_no, '')
                  WHEN '' THEN 0
                  ELSE 1
             END) AS newNum ,
         SUM(CASE ISNULL(faq_no, '')
                  WHEN '' THEN 1
                  ELSE 0
             END) AS noHandleNum ,
         COUNT(*) AS totalNum
FROM     QAToSissKC ,
         qauser
WHERE    status = '1'
         AND back_date >= '{0}'
         AND back_date <= '{1}'
         AND back_user = userid
GROUP BY name;

";
        internal static string zskzltjQuery = @"
SELECT   '平台转入' AS project ,
         name ,
         COUNT(*) AS totalNum ,
         SUM(CASE ISNULL(faq_no, '')
                  WHEN '' THEN 0
                  ELSE 1
             END) AS validNum
FROM     iss.QAToSissKC ,
         iss.qauser
WHERE    status = '1'
         AND CONVERT(CHAR(10), back_date, 121) >= '{0}'
         AND CONVERT(CHAR(10), back_date, 121) <= '{1}'
         AND iss.qauser.userid = iss.QAToSissKC.back_user
GROUP BY name
UNION
SELECT   '我要分享' AS project ,
         name ,
         COUNT(*) AS totalNum ,
         SUM(CASE ISNULL(approve_kcid, '')
                  WHEN '' THEN 0
                  ELSE 1
             END) AS validNum
FROM     iss.SissKCUserQA ,
         iss.qauser
WHERE    iss.qauser.userid = iss.SissKCUserQA.back_user
         AND CONVERT(CHAR(10), back_date, 121) >= '{0}'
         AND CONVERT(CHAR(10), back_date, 121) <= '{1}'
         AND ISNULL(back_user, '') <> ''
GROUP BY name
ORDER BY name;
";
    }
}
