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
            return QyxnSqlQuery(Config.connectionString, QueryStrings.qyxnQuery, Config.beginTime, Config.endTime, dept);
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

        private static DataTable QyxnSqlQuery(string connectionString, string sqlString, string beginTime, string endTime, string dept)
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
