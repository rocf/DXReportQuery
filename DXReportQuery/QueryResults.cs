﻿using System;
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
            Config Config = new Config();
            return SqlQuery(Config.connectionString, QueryStrings.djwtQuery, Config.beginTime, Config.endTime);               
        }

        public static DataTable ZtgblQuery()
        {
            Config Config = new Config();
            return SqlQuery(Config.connectionString, QueryStrings.ztgblQuery, Config.beginTime, Config.endTime);
        }

        public static DataTable VIPWtgblQuery()
        {
            Config Config = new Config();
            return SqlQuery(Config.connectionString, QueryStrings.vipWtgblQuery, Config.beginTime, Config.endTime);
        }

        public static DataTable ClzWtclQuery()
        {
            Config Config = new Config();
            return SqlQuery(Config.connectionString, QueryStrings.clzWtclQuery, Config.beginTime, Config.endTime);
        }

        public static DataTable QyxnQuery(string dept)
        {
            Config Config = new Config();
            return DeptSqlQuery(Config.connectionString, QueryStrings.qyxnQuery, Config.beginTime, Config.endTime, dept);
        }

        public static DataTable VIPGblQuery(string dept)
        {
            Config Config = new Config();
            return DeptSqlQuery(Config.connectionString, QueryStrings.vipGblQuery, Config.beginTime, Config.endTime, dept);
        }

        public static DataTable QybbQuery(string dept)
        {
            Config Config = new Config();
            return DeptSqlQuery(Config.connectionString, QueryStrings.qybbQuery, Config.beginTime, Config.endTime, dept);
        }

        public static DataTable GrxnQuery(string dept)
        {
            Config Config = new Config();
            return DeptSqlQuery(Config.connectionString, QueryStrings.grxnQuery, Config.beginTime, Config.endTime, dept);
        }
        public static DataTable VIPKhjlGrxnQuery(string dept)
        {
            Config Config = new Config();
            return DeptSqlQuery(Config.connectionString, QueryStrings.vipKhjlGrxnQuery, Config.beginTime, Config.endTime, dept);
        }

        public static DataTable DlsyjQuery(string dept)
        {
            Config Config = new Config();
            return DeptSqlQuery(Config.connectionString, QueryStrings.dlsyjQuery, Config.beginTime, Config.endTime, dept);
        }

        public static DataTable VIPDlsyjQuery(string dept)
        {
            Config Config = new Config();
            return DeptSqlQuery(Config.connectionString, QueryStrings.vipDlsyjQuery, Config.beginTime, Config.endTime, dept);
        }

        public static DataTable WtyjQuery()
        {
            Config Config = new Config();
            return SqlQuery(Config.connectionString, QueryStrings.wtyjQuery, Config.beginTime, Config.endTime);
        }

        public static DataTable WtxqzbQuery()
        {
            Config Config = new Config();
            return SqlQuery(Config.connectionString, QueryStrings.wtxqzblQuery, Config.beginTime, Config.endTime);
        }

        public static DataTable ZzsktjQuery()
        {
            Config Config = new Config();
            return SqlQuery(Config.connectionString, QueryStrings.zzsktjQuery, Config.beginTime, Config.endTime);
        }

        public static DataTable ZskclsltjQuery()
        {
            Config Config = new Config();
            return SqlQuery(Config.connectionString, QueryStrings.zskclsltjQuery, Config.beginTime, Config.endTime);
        }

        public static DataTable WtxqzblQuery()
        {
            Config Config = new Config();
            return SqlQuery(Config.connectionString, QueryStrings.wtxqzblQuery, Config.beginTime, Config.endTime);
        }
        public static DataTable ZskzltjQuery()
        {
            Config Config = new Config();
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
