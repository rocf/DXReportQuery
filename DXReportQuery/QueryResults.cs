using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using DevExpress.XtraSpreadsheet.Commands;

namespace DXReportQuery
{
     class QueryResults
    {
 
        public static DataTable DjwtQuery()
        {
            return SqlQuery(Config.connectionString, QueryStrings.djwtQuery, Config.beginTime, Config.endTime);               
        }

        public static DataTable ZtgblQuery()
        {
            return SqlQuery(Config.connectionString, QueryStrings.ztgblQuery, Config.beginTime, Config.endTime);
        }

        public static DataTable VIPWtgblQuery()
        {
            return SqlQuery(Config.connectionString, QueryStrings.vipWtgblQuery, Config.beginTime, Config.endTime);
        }

        public static DataTable ClzWtclQuery()
        {
            return SqlQuery(Config.connectionString, QueryStrings.clzWtclQuery, Config.beginTime, Config.endTime);
        }

        public static DataTable QyxnQuery(string dept)
        {
            return DeptSqlQuery(Config.connectionString, QueryStrings.qyxnQuery, Config.beginTime, Config.endTime, dept);
        }

        public static DataTable VIPGblQuery(string dept)
        {
            return DeptSqlQuery(Config.connectionString, QueryStrings.vipGblQuery, Config.beginTime, Config.endTime, dept);
        }

        public static DataTable QybbQuery(string dept)
        {
            return DeptSqlQuery(Config.connectionString, QueryStrings.qybbQuery, Config.beginTime, Config.endTime, dept);
        }

        public static DataTable GrxnQuery(string dept)
        {
            return DeptSqlQuery(Config.connectionString, QueryStrings.grxnQuery, Config.beginTime, Config.endTime, dept);
        }
        public static DataTable VIPKhjlGrxnQuery(string dept)
        {
            return DeptSqlQuery(Config.connectionString, QueryStrings.vipKhjlGrxnQuery, Config.beginTime, Config.endTime, dept);
        }

        public static DataTable DlsyjQuery(string dept)
        {
            return DeptSqlQuery(Config.connectionString, QueryStrings.dlsyjQuery, Config.beginTime, Config.endTime, dept);
        }

        public static DataTable VIPDlsyjQuery(string dept)
        {
            return DeptSqlQuery(Config.connectionString, QueryStrings.vipDlsyjQuery, Config.beginTime, Config.endTime, dept);
        }

        public static DataTable WtyjQuery()
        {
            return SqlQuery(Config.connectionString, QueryStrings.wtyjQuery, Config.beginTime, Config.endTime);
        }

        public static DataTable WtxqzbQuery()
        {
            return SqlQuery(Config.connectionString, QueryStrings.wtxqzbQuery, Config.beginTime, Config.endTime);
        }

        public static DataTable ZzsktjQuery()
        {
            return SqlQuery(Config.connectionString, QueryStrings.zzsktjQuery, Config.beginTime, Config.endTime);
        }

        public static DataTable ZskclsltjQuery()
        {
            return SqlQuery(Config.connectionString, QueryStrings.zskclsltjQuery, Config.beginTime, Config.endTime);
        }

        public static DataTable ZskzltjQuery()
        {
            return SqlQuery(Config.connectionString, QueryStrings.zskzltjQuery, Config.beginTime, Config.endTime);
        }
        private static DataTable SqlQuery(string connectionString, string sqlString, string beginTime, string endTime)
        {
            using (SqlConnection sqlConnection = new SqlConnection())
            {
                sqlConnection.ConnectionString = connectionString;

                using (SqlCommand sqlCommand = new SqlCommand())
                {
                    sqlConnection.Open();
                    sqlCommand.CommandText =string.Format(sqlString, beginTime, endTime);
                    sqlCommand.Connection = sqlConnection;
                    SqlDataReader sqlDataReader = sqlCommand.ExecuteReader();
                    DataTable dataTable = new DataTable();
                    dataTable.Load(sqlDataReader);

                    return dataTable;
                }
            }
        }

        private static DataTable DeptSqlQuery(string connectionString, string sqlString, string beginTime, string endTime, string dept)
        {
            using (SqlConnection sqlConnection = new SqlConnection())
            {
                sqlConnection.ConnectionString = connectionString;

                using (SqlCommand sqlCommand = new SqlCommand())
                {
                    sqlConnection.Open();
                    sqlCommand.CommandText = string.Format(sqlString, beginTime, endTime, dept);
                    sqlCommand.Connection = sqlConnection;
                    SqlDataReader sqlDataReader = sqlCommand.ExecuteReader();
                    DataTable dataTable = new DataTable();
                    dataTable.Load(sqlDataReader);

                    return dataTable;
                }
            }
        }
    }
}
