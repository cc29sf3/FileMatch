using System;
using System.Collections;
using System.Data;
using System.Data.SqlClient;

namespace Utility.Dao
{
    /// <summary>
    /// 针对SQL数据库的操作
    /// 名称：SqlHelper
    /// 作者：任大可
    /// 创建日期：2014-04-13
    /// </summary>
    public sealed class SqlHelper
    {
        private string connectionString;

        /// <summary>
        /// 构造函数
        /// </summary>
        /// <param name="strLoginFile">数据文件信息</param>
        public SqlHelper(string strLoginFile)
        { connectionString = strLoginFile; }

        /// <summary>
        /// SQL参数缓存
        /// </summary>
        private Hashtable parmCache = Hashtable.Synchronized(new Hashtable());

        /// <summary>
        /// 缓存SQL参数
        /// </summary>
        /// <param name="cacheKey">要缓存的参数的键</param>
        /// <param name="commandParameters">要缓存的SQL参数值</param>
        public void CacheParameters(string cacheKey, params SqlParameter[] commandParameters)
        { parmCache[cacheKey] = commandParameters; }

        /// <summary>
        /// 利用存储过程和指定参数来对记录执行增删改操作
        /// </summary>
        /// <param name="UspName">存储过程名字</param>
        /// <param name="param">存储过程参数</param>
        /// <returns>返回受影响的记录条数</returns>
        public int DeleteByWhere(string UspName, SqlParameter[] param)
        {
            SqlCommand cmd = new SqlCommand();
            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                PrepareCommand(cmd, connection, null, CommandType.StoredProcedure, UspName, param);
                int num = cmd.ExecuteNonQuery();
                cmd.Parameters.Clear();
                return num;
            }
        }

        /// <summary>
        /// 利用SQL语句来对记录执行增删改操作
        /// </summary>
        /// <param name="cmdText">要执行的增删改的SQL语句</param>
        /// <returns>返回受影响的行数</returns>
        public int ExecuteNonQuery(string cmdText)
        { return ExecuteNonQuery(connectionString, CommandType.Text, cmdText, null); }

        /// <summary>
        /// 对记录执行增删改操作
        /// </summary>
        /// <param name="cmdType">要执行的SQL命令的类型</param>
        /// <param name="cmdText">要执行的SQL命令文本</param>
        /// <returns>受影响的行数</returns>
        public int ExecuteNonQuery(CommandType cmdType, string cmdText)
        { return ExecuteNonQuery(connectionString, cmdType, cmdText, null); }

        /// <summary>
        /// 对记录执行增删改操作
        /// </summary>
        /// <param name="cmdType">要执行的SQL命令的类型</param>
        /// <param name="cmdText">要执行的SQL命令文本</param>
        /// <param name="parameters">参数列表</param>
        /// <returns>受影响的行数</returns>
        public int ExecuteNonQuery(CommandType cmdType, string cmdText, params SqlParameter[] parameters)
        { return ExecuteNonQuery(connectionString, cmdType, cmdText, parameters); }

        /// <summary>
        /// 对记录执行增删改操作
        /// </summary>
        /// <param name="connection">数据库连接</param>
        /// <param name="cmdType">要执行的SQL命令的类型</param>
        /// <param name="cmdText">要执行的SQL命令文本</param>
        /// <param name="parameters">参数列表</param>
        /// <returns>受影响的行数</returns>
        public int ExecuteNonQuery(SqlConnection connection, CommandType cmdType, string cmdText, params SqlParameter[] parameters)
        {
            SqlCommand cmd = new SqlCommand();
            int num = 0;
            PrepareCommand(cmd, connection, null, cmdType, cmdText, parameters);
            using (connection)
            {
                num = cmd.ExecuteNonQuery();
                cmd.Parameters.Clear();
            }
            return num;
        }

        /// <summary>
        /// 对记录执行增删改操作
        /// </summary>
        /// <param name="trans">数据库事务实例</param>
        /// <param name="cmdType">要执行的SQL命令的类型</param>
        /// <param name="cmdText">要执行的SQL命令文本</param>
        /// <param name="parameters">参数列表</param>
        /// <returns>受影响的行数</returns>
        public int ExecuteNonQuery(SqlTransaction trans, CommandType cmdType, string cmdText, params SqlParameter[] parameters)
        {
            SqlCommand cmd = new SqlCommand();
            int num = -1;
            PrepareCommand(cmd, trans.Connection, trans, cmdType, cmdText, parameters);
            using (trans.Connection)
            {
                num = cmd.ExecuteNonQuery();
                cmd.Parameters.Clear();
            }
            return num;
        }
       
        /// <summary>
        /// 对记录执行增删改操作
        /// </summary>
        /// <param name="connectionString">数据库连接字符串</param>
        /// <param name="cmdType">要执行的SQL命令的类型</param>
        /// <param name="cmdText">要执行的SQL命令文本</param>
        /// <param name="parameters">参数列表</param>
        /// <returns>受影响的行数</returns>
        public int ExecuteNonQuery(string connectionString, CommandType cmdType, string cmdText, params SqlParameter[] parameters)
        {
            SqlCommand cmd = new SqlCommand(); int num = -1;
            using (SqlConnection connection = new SqlConnection(connectionString))
                try
                {
                    PrepareCommand(cmd, connection, null, cmdType, cmdText, parameters);
                    num = cmd.ExecuteNonQuery();
                    cmd.Parameters.Clear();
                    cmd.Dispose();
                }
                catch (OutOfMemoryException e)
                { Export.Outputerror(e.Message, false); }
                catch (Exception e)
                { Export.Outputerror(e.Message); }
            return num;
        }

        /// <summary>
        /// 返回DataReadera实例
        /// </summary>
        /// <param name="cmdText">查询SQL命令</param>
        /// <returns>DataReadera实例</returns>
        public SqlDataReader ExecuteReader(string cmdText)
        { return ExecuteReader(connectionString, CommandType.Text, cmdText, null); }

        /// <summary>
        /// 返回DataReadera实例
        /// </summary>
        /// <param name="cmdType">查询SQL命令的类型，如普通SQL语句或存储过程</param>
        /// <param name="cmdText">查询SQL命令</param>
        /// <returns>DataReadera实例</returns>
        public SqlDataReader ExecuteReader(CommandType cmdType, string cmdText)
        { return ExecuteReader(connectionString, cmdType, cmdText, null); }

        /// <summary>
        /// 返回DataReadera实例
        /// </summary>
        /// <param name="cmdType">查询SQL命令的类型，如普通SQL语句或存储过程</param>
        /// <param name="cmdText">查询SQL命令</param>
        /// <param name="parameters">查询SQL命令所需要的参数</param>
        /// <returns>DataReadera实例</returns>
        public SqlDataReader ExecuteReader(CommandType cmdType, string cmdText, params SqlParameter[] parameters)
        { return ExecuteReader(connectionString, cmdType, cmdText, parameters); }

        /// <summary>
        /// 返回DataReadera实例
        /// </summary>
        /// <param name="connectionString">数据库连接字符串</param>
        /// <param name="cmdType">查询SQL命令的类型，如普通SQL语句或存储过程</param>
        /// <param name="cmdText">查询SQL命令</param>
        /// <param name="parameters">查询SQL命令所需要的参数</param>
        /// <returns>DataReadera实例</returns>
        public SqlDataReader ExecuteReader(string connectionString, CommandType cmdType, string cmdText, params SqlParameter[] parameters)
        {
            SqlCommand cmd = new SqlCommand();
            SqlConnection conn = new SqlConnection(connectionString);
            PrepareCommand(cmd, conn, null, cmdType, cmdText, parameters);
            SqlDataReader reader = cmd.ExecuteReader(CommandBehavior.CloseConnection);
            cmd.Parameters.Clear();
            conn.Close();
            return reader;
        }

        public SqlDataReader ExecuteReader(SqlConnection conn, string cmdText)
        {
            SqlCommand cmd = new SqlCommand();
            PrepareCommand(cmd, conn, null, CommandType.Text, cmdText, null);
            SqlDataReader reader = cmd.ExecuteReader(CommandBehavior.CloseConnection);
            cmd.Parameters.Clear();
            return reader;
        }

        /// <summary>
        /// 执行返回结果集的查询
        /// </summary>
        /// <param name="procedureName">存储过程名字</param>
        /// <param name="parameters">存储过程用到的参数</param>
        /// <returns>返回结果集</returns>
        public DataSet GetDataSet(string procedureName, params SqlParameter[] parameters)
        {
            SqlCommand cmd = new SqlCommand();
            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                PrepareCommand(cmd, connection, null, CommandType.StoredProcedure, procedureName, parameters);
                SqlDataAdapter adapter = new SqlDataAdapter(cmd);
                DataSet dataSet = new DataSet();
                adapter.Fill(dataSet);
                cmd.Parameters.Clear();
                return dataSet;
            }
        }

        /// <summary>
        /// 执行返回结果集的查询
        /// </summary>
        /// <param name="procedureName">存储过程名称</param>
        /// <param name="parameters">查询SQL命令所需要的参数</param>
        /// <returns>返回结果集</returns>
        public DataTable GetDataTable(string procedureName, params SqlParameter[] parameters)
        { return ExecuteTable(CommandType.StoredProcedure, procedureName, parameters); }

        /// <summary>
        /// 执行返回结果集的查询
        /// </summary>
        /// <param name="cmdText">查询SQL命令</param>
        /// <returns>返回结果集</returns>
        public DataTable ExecuteTable(string cmdText)
        {
            DataTable dataTable = new DataTable();

            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                SqlCommand cmd = new SqlCommand();
                PrepareCommand(cmd, connection, null, CommandType.Text, cmdText, null);
                SqlDataAdapter adapter = new SqlDataAdapter(cmd);
                try
                { adapter.Fill(dataTable); }
                catch (OutOfMemoryException e)
                { Export.Outputerror(e.Message, false); }
                catch (Exception e)
                { Export.Outputerror(e.Message); }
                cmd.Parameters.Clear();
            }
            return dataTable;
        }

        /// <summary>
        /// 执行返回结果集的查询
        /// </summary>
        /// <param name="cmdType">查询SQL命令的类型，如普通SQL语句或存储过程</param>
        /// <param name="cmdText">查询SQL命令</param>
        /// <param name="parameters">查询SQL命令所需要的参数</param>
        /// <returns>返回结果集</returns>
        public DataTable ExecuteTable(CommandType cmdType, string cmdText, params SqlParameter[] parameters)
        {
            SqlCommand cmd = new SqlCommand();
            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                PrepareCommand(cmd, connection, null, cmdType, cmdText, parameters);
                SqlDataAdapter adapter = new SqlDataAdapter(cmd);
                DataTable dataTable = new DataTable();
                adapter.Fill(dataTable);
                cmd.Parameters.Clear();
                return dataTable;
            }
        }

        /// <summary>
        /// 执行结果集为单行单列的查询
        /// </summary>
        /// <param name="cmdText">查询SQL命令</param>
        /// <returns>返回查询结果的第一行第一列</returns>
        public object ExecuteScalar(string cmdText)
        { return ExecuteScalar(connectionString, CommandType.Text, cmdText, null); }

        /// <summary>
        /// 执行结果集为单行单列的查询
        /// </summary>
        /// <param name="cmdType">查询SQL命令的类型，如普通SQL语句或存储过程</param>
        /// <param name="cmdText">查询SQL命令</param>
        /// <returns>返回查询结果的第一行第一列</returns>
        public object ExecuteScalar(CommandType cmdType, string cmdText)
        { return ExecuteScalar(connectionString, cmdType, cmdText, null); }

        /// <summary>
        /// 执行结果集为单行单列的查询
        /// </summary>
        /// <param name="connectionString">数据库连接字符串</param>
        /// <param name="cmdType">查询SQL命令的类型，如普通SQL语句或存储过程</param>
        /// <param name="cmdText">查询SQL命令</param>
        /// <param name="parameters">查询SQL命令所需要的参数</param>
        /// <returns>返回查询结果的第一行第一列</returns>
        public object ExecuteScalar(CommandType cmdType, string cmdText, params SqlParameter[] parameters)
        { return ExecuteScalar(connectionString, cmdType, cmdText, parameters); }

        /// <summary>
        /// 执行结果集为单行单列的查询
        /// </summary>
        /// <param name="connection">数据库连接实例</param>
        /// <param name="cmdType">查询SQL命令的类型，如普通SQL语句或存储过程</param>
        /// <param name="cmdText">查询SQL命令</param>
        /// <param name="parameters">查询SQL命令所需要的参数</param>
        /// <returns>返回查询结果的第一行第一列</returns>
        public object ExecuteScalar(SqlConnection connection, CommandType cmdType, string cmdText, params SqlParameter[] parameters)
        {
            SqlCommand cmd = new SqlCommand();
            PrepareCommand(cmd, connection, null, cmdType, cmdText, parameters);
            object obj2 = cmd.ExecuteScalar();
            cmd.Parameters.Clear();
            return obj2;
        }

        /// <summary>
        /// 执行结果集为单行单列的查询
        /// </summary>
        /// <param name="connectionString">数据库连接字符串</param>
        /// <param name="cmdType">查询SQL命令的类型，如普通SQL语句或存储过程</param>
        /// <param name="cmdText">查询SQL命令</param>
        /// <param name="parameters">查询SQL命令所需要的参数</param>
        /// <returns>返回查询结果的第一行第一列</returns>
        public object ExecuteScalar(string connectionString, CommandType cmdType, string cmdText, params SqlParameter[] parameters)
        {
            SqlCommand cmd = new SqlCommand();
            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                PrepareCommand(cmd, connection, null, cmdType, cmdText, parameters);
                object obj2 = cmd.ExecuteScalar();
                cmd.Parameters.Clear();
                return obj2;
            }
        }

        /// <summary>
        ///
        /// </summary>
        /// <param name="cacheKey"></param>
        /// <returns></returns>
        public SqlParameter[] GetCachedParameters(string cacheKey)
        {
            SqlParameter[] cachedParms = (SqlParameter[])parmCache[cacheKey];
            if (cachedParms == null)
                return null;
            SqlParameter[] clonedParms = new SqlParameter[cachedParms.Length];
            for (int i = 0, j = cachedParms.Length; i < j; i++)
                clonedParms[i] = (SqlParameter)((ICloneable)cachedParms[i]).Clone();
            return clonedParms;
        }

        /// <summary>
        /// 取得Commad的值
        /// </summary>
        /// <param name="cmdText">存储过程</param>
        /// <param name="name">要取值的SQL命令参数</param>
        /// <param name="parameters">SQL命令参数</param>
        /// <returns>object</returns>
        public object GetParameterValue(string cmdText, string name, params SqlParameter[] parameters)
        {
            SqlCommand cmd = new SqlCommand();
            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                PrepareCommand(cmd, connection, null, CommandType.StoredProcedure, cmdText, parameters);
                cmd.ExecuteNonQuery();
                object value = cmd.Parameters[name].Value;
                cmd.Parameters.Clear();
                return value;
            }
        }

        /// <summary>
        ///  插入表
        /// </summary>
        /// <param name="Item">表</param>
        /// <param name="dt">DataTable实例</param>
        /// <param name="listed">列数组</param>
        /// <returns>bool</returns>
        public bool InsertTable(string Item, DataTable dt, string[] listed)
        {
            bool bo = false;
            using (SqlConnection conn = new SqlConnection(connectionString))
            using (SqlBulkCopy bulk = new SqlBulkCopy(conn))
            {
                foreach (string Row in listed)
                    bulk.ColumnMappings.Add(Row, Row);
                bulk.DestinationTableName = Item;
                bulk.NotifyAfter = 100;
                bulk.BulkCopyTimeout = 100;
                bulk.BatchSize = 10;
                try
                {
                    conn.Open();
                    bulk.WriteToServer(dt);
                    bo = true;
                }
                catch (OutOfMemoryException e)
                { Export.Outputerror(e.Message, false); }
                catch (Exception e)
                { Export.Outputerror(e.Message); }
                conn.Close();
            }
            return bo;
        }

        /// <summary>
        /// 给SqlCommand实例指定参数信息
        /// </summary>
        /// <param name="cmd">SqlCommand实例</param>
        /// <param name="conn">SqlConnection实例</param>
        /// <param name="trans">数据库事务实例</param>
        /// <param name="cmdType">SQL命令类型</param>
        /// <param name="cmdText">SQL命令</param>
        /// <param name="cmdParms">SQL命令参数</param>
        private void PrepareCommand(SqlCommand cmd, SqlConnection conn, SqlTransaction trans, CommandType cmdType, string cmdText, SqlParameter[] cmdParms)
        {
            try
            {
                if (conn.State != ConnectionState.Open)
                    conn.Open();
                cmd.Connection = conn;
                cmd.CommandText = cmdText;
                if (trans != null)
                    cmd.Transaction = trans;
                cmd.CommandType = cmdType;
                if (cmdParms != null)
                    foreach (SqlParameter parameter in cmdParms)
                        cmd.Parameters.Add(parameter);
            }
            catch (OutOfMemoryException e)
            { Export.Outputerror(e.Message, false); }
            catch (Exception e)
            { Export.Outputerror(e.Message); };
        }
    }
}