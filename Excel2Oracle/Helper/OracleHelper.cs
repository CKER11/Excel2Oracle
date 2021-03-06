﻿using System;
using System.Data;
using System.Collections.Generic;
using System.Configuration;
using System.Data.OracleClient;
using System.Text;
using System.IO;
using System.Windows.Forms;
using System.Xml;

namespace Excel2Oracle
{
    class OracleHelper
    {
        //BGService.BusinessServiceClient sqlClient = new BGService.BusinessServiceClient();
        //public int ConnectionTest(string connStr)
        //{
        //    using (var conn = new OracleConnection(connStr))
        //    {
        //        conn.Open();
        //        DataSet ds = new DataSet();
        //        string sql = "select 1 from dual";
        //        OracleDataAdapter oda = new OracleDataAdapter(sql, conn);
        //        oda.Fill(ds);
        //        DataTable dt = ds.Tables[0];
        //        if (dt.Rows.Count > 0)
        //        {
        //            return 1;
        //        }
        //        return 0;
        //    }
        //}

        public static int ConnectionTest(string connStr)
        {
            using (var conn = new OracleConnection(connStr))
            {
                conn.Open();
                DataSet ds = new DataSet();
                string sql = "select 1 from dual";
                OracleDataAdapter oda = new OracleDataAdapter(sql, conn);
                oda.Fill(ds);
                DataTable dt = ds.Tables[0];
                if (dt.Rows.Count > 0)
                {
                    return 1;
                }
                return 0;
            }
        }
        /// <summary>  
        /// 执行数据库增删改操作,返回受影响的行数
        /// </summary>  
        /// <param name="connectionString">数据库连接字符串</param>
        /// <param name="cmdType">命令的类型</param>
        /// <param name="cmdText">Oracle存储过程名称或PL/SQL命令</param>  
        /// <param name="cmdParms">命令参数集合</param>  
        /// <returns>当前查询操作返回的OracleDataReader类型的内存结果集</returns>  
        public static int ExecuteNonQuery(string connectionString, CommandType cmdType, string cmdText, params OracleParameter[] cmdParms)
        {
            OracleCommand cmd = new OracleCommand();
            OracleConnection conn = new OracleConnection(connectionString);
            try
            {
                PrepareCommand(cmd, conn, null, cmdType, cmdText, cmdParms);
                cmd.Parameters.Clear();
                return cmd.ExecuteNonQuery();
            }
            catch
            {
                cmd.Dispose();
                conn.Close();
                throw;
            }
        }

        /// <summary>  
        /// 执行数据库查询操作,返回OracleDataReader类型的内存结果集  
        /// </summary>  
        /// <param name="connectionString">数据库连接字符串</param>
        /// <param name="cmdType">命令的类型</param>
        /// <param name="cmdText">Oracle存储过程名称或PL/SQL命令</param>  
        /// <param name="cmdParms">命令参数集合</param>  
        /// <returns>当前查询操作返回的OracleDataReader类型的内存结果集</returns>  
        public static OracleDataReader ExecuteReader(string connectionString, CommandType cmdType, string cmdText, params OracleParameter[] cmdParms)
        {
            OracleCommand cmd = new OracleCommand();
            OracleConnection conn = new OracleConnection(connectionString);
            try
            {
                PrepareCommand(cmd, conn, null, cmdType, cmdText, cmdParms);
                OracleDataReader reader = cmd.ExecuteReader(CommandBehavior.CloseConnection);
                cmd.Parameters.Clear();
                return reader;
            }
            catch
            {
                cmd.Dispose();
                conn.Close();
                throw;
            }
        }

        /// <summary>  
        /// 执行数据库查询操作,返回DataSet类型的结果集  
        /// </summary>  
        /// <param name="connectionString">数据库连接字符串</param>
        /// <param name="cmdType">命令的类型</param>
        /// <param name="cmdText">Oracle存储过程名称或PL/SQL命令</param>  
        /// <param name="cmdParms">命令参数集合</param>  
        /// <returns>当前查询操作返回的DataSet类型的结果集</returns>  
        public static DataSet ExecuteDataSet(string connectionString, CommandType cmdType, string cmdText, params OracleParameter[] cmdParms)
        {
            OracleCommand cmd = new OracleCommand();
            OracleConnection conn = new OracleConnection(connectionString);
            DataSet ds = null;
            try
            {
                PrepareCommand(cmd, conn, null, cmdType, cmdText, cmdParms);
                OracleDataAdapter adapter = new OracleDataAdapter();
                adapter.SelectCommand = cmd;
                ds = new DataSet();
                adapter.Fill(ds);
                cmd.Parameters.Clear();
            }
            catch
            {
                throw;
            }
            finally
            {
                cmd.Dispose();
                conn.Close();
                conn.Dispose();
            }

            return ds;
        }

        /// <summary>  
        /// 执行数据库查询操作,返回DataTable类型的结果集  
        /// </summary>  
        /// <param name="connectionString">数据库连接字符串</param>
        /// <param name="cmdType">命令的类型</param>
        /// <param name="cmdText">Oracle存储过程名称或PL/SQL命令</param>  
        /// <param name="cmdParms">命令参数集合</param>  
        /// <returns>当前查询操作返回的DataTable类型的结果集</returns>  
        public static DataTable ExecuteDataTable(string connectionString, CommandType cmdType, string cmdText, params OracleParameter[] cmdParms)
        {
            OracleCommand cmd = new OracleCommand();
            OracleConnection conn = new OracleConnection(connectionString);
            DataTable dt = null;

            try
            {
                PrepareCommand(cmd, conn, null, cmdType, cmdText, cmdParms);
                OracleDataAdapter adapter = new OracleDataAdapter();
                adapter.SelectCommand = cmd;
                dt = new DataTable();
                adapter.Fill(dt);
                cmd.Parameters.Clear();
            }
            catch
            {
                throw;
            }
            finally
            {
                cmd.Dispose();
                conn.Close();
                conn.Dispose();
            }

            return dt;
        }

        /// <summary>  
        /// 执行数据库查询操作,返回结果集中位于第一行第一列的Object类型的值  
        /// </summary>  
        /// <param name="connectionString">数据库连接字符串</param>
        /// <param name="cmdType">命令的类型</param>
        /// <param name="cmdText">Oracle存储过程名称或PL/SQL命令</param>  
        /// <param name="cmdParms">命令参数集合</param>  
        /// <returns>当前查询操作返回的结果集中位于第一行第一列的Object类型的值</returns>  
        public static object ExecuteScalar(string connectionString, CommandType cmdType, string cmdText, params OracleParameter[] cmdParms)
        {
            OracleCommand cmd = new OracleCommand();
            OracleConnection conn = new OracleConnection(connectionString);
            object result = null;
            try
            {
                PrepareCommand(cmd, conn, null, cmdType, cmdText, cmdParms);
                result = cmd.ExecuteScalar();
                cmd.Parameters.Clear();
            }
            catch
            {
                throw;
            }
            finally
            {
                cmd.Dispose();
                conn.Close();
                conn.Dispose();
            }

            return result;
        }



        /// <summary>  
        /// 执行数据库命令前的准备工作  
        /// </summary>  
        /// <param name="cmd">Command对象</param>  
        /// <param name="conn">数据库连接对象</param>  
        /// <param name="trans">事务对象</param>  
        /// <param name="cmdType">Command类型</param>  
        /// <param name="cmdText">Oracle存储过程名称或PL/SQL命令</param>  
        /// <param name="cmdParms">命令参数集合</param>  
        private static void PrepareCommand(OracleCommand cmd, OracleConnection conn, OracleTransaction trans, CommandType cmdType, string cmdText, OracleParameter[] cmdParms)
        {
            if (conn.State != ConnectionState.Open)
                conn.Open();

            cmd.Connection = conn;
            cmd.CommandText = cmdText;

            if (trans != null)
                cmd.Transaction = trans;

            cmd.CommandType = cmdType;

            if (cmdParms != null)
            {
                foreach (OracleParameter parm in cmdParms)
                    cmd.Parameters.Add(parm);
            }
        }

        /// <summary>  
        /// 将.NET日期时间类型转化为Oracle兼容的日期时间格式字符串  
        /// </summary>  
        /// <param name="date">.NET日期时间类型对象</param>  
        /// <returns>Oracle兼容的日期时间格式字符串（如该字符串：TO_DATE('2007-12-1','YYYY-MM-DD')）</returns>  
        public static string GetOracleDateFormat(DateTime date)
        {
            return "TO_DATE('" + date.ToString("yyyy-M-dd") + "','YYYY-MM-DD')";
        }

        /// <summary>  
        /// 将.NET日期时间类型转化为Oracle兼容的日期格式字符串  
        /// </summary>  
        /// <param name="date">.NET日期时间类型对象</param>  
        /// <param name="format">Oracle日期时间类型格式化限定符</param>  
        /// <returns>Oracle兼容的日期时间格式字符串（如该字符串：TO_DATE('2007-12-1','YYYY-MM-DD')）</returns>  
        public static string GetOracleDateFormat(DateTime date, string format)
        {
            if (format == null || format.Trim() == "") format = "YYYY-MM-DD";
            return "TO_DATE('" + date.ToString("yyyy-M-dd") + "','" + format + "')";
        }

        /// <summary>  
        /// 将指定的关键字处理为模糊查询时的合法参数值  
        /// </summary>  
        /// <param name="source">待处理的查询关键字</param>  
        /// <returns>过滤后的查询关键字</returns>  
        public static string HandleLikeKey(string source)
        {
            if (source == null || source.Trim() == "") return null;

            source = source.Replace("[", "[]]");
            source = source.Replace("_", "[_]");
            source = source.Replace("%", "[%]");

            return ("%" + source + "%");
        }

    }

}