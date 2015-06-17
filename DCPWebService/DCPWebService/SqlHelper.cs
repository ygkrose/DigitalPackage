using System;
using System.Collections.Generic;
using System.Text;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Collections;
using System.Xml;


namespace DCPWebService
{
    public abstract class SqlHelper
    {
        public static string ConnectionStringLocalTransaction
        {
            get { return ConfigurationSettings.AppSettings["SqlConnStr"]; }
        }

        // Hashtable to store cached parameters
        private static Hashtable parmCache = Hashtable.Synchronized(new Hashtable());


        /// <summary>
        /// ����@�� SqlCommand (�^�Ǩ��v�T����ƦC�ƥ�)  
        /// �ϥ�parameters�����ѼƱ���
        /// </summary>
        /// <remarks>
        /// e.g.:  
        ///  int result = ExecuteNonQuery(connString, CommandType.StoredProcedure, "PublishOrders", new SqlParameter("@prodid", 24));
        /// </remarks>
        /// <param name="connectionString">�@�ӦX�k��SqlConnection��connection string</param>
        /// <param name="commandType">�R�O�Φ�(stored procedure, text, etc.)</param>
        /// <param name="commandText">stored procedure�W�٩�T-SQL�R�O</param>
        /// <param name="commandParameters">�����檺�R�O��SqlParamters�}�C</param>
        /// <returns>�^�Ǧ��R�O���v�T����ƦC�ƥ�</returns>
        public static int ExecuteNonQuery(string connectionString, CommandType cmdType, string cmdText, params SqlParameter[] commandParameters)
        {

            SqlCommand cmd = new SqlCommand();

            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                PrepareCommand(cmd, conn, null, cmdType, cmdText, commandParameters);
                int val = cmd.ExecuteNonQuery();
                cmd.Parameters.Clear();
                return val;
            }
        }

        /// <summary>
        /// ����@�� SqlCommand (���^�ǵ��G) using an existing SQL Transaction 
        /// �ϥ�parameters�����ѼƱ���
        /// </summary>
        /// <remarks>
        /// e.g.:  
        ///  int result = ExecuteNonQuery(connString, CommandType.StoredProcedure, "PublishOrders", new SqlParameter("@prodid", 24));
        /// </remarks>
        /// <param name="trans">an existing sql transaction</param>
        /// <param name="commandType">�R�O�Φ�(stored procedure, text, etc.)</param>
        /// <param name="commandText">stored procedure�W�٩�T-SQL�R�O</param>
        /// <param name="commandParameters">�����檺�R�O��SqlParamters�}</param>
        /// <returns>�^�Ǧ��R�O���v�T����ƦC�ƥ�</returns>
        public static int ExecuteNonQuery(SqlTransaction trans, CommandType cmdType, string cmdText, params SqlParameter[] commandParameters)
        {
            SqlCommand cmd = new SqlCommand();
            PrepareCommand(cmd, trans.Connection, trans, cmdType, cmdText, commandParameters);
            int val = cmd.ExecuteNonQuery();
            cmd.Parameters.Clear();
            return val;
        }

        /// <summary>
        /// ����@�� SqlCommand (�^�Ǩ��v�T����ƶ��X)
        /// �ϥ�parameters�����ѼƱ���
        /// </summary>
        /// <remarks>
        /// e.g.:  
        ///  SqlDataReader r = ExecuteReader(connString, CommandType.StoredProcedure, "PublishOrders", new SqlParameter("@prodid", 24));
        /// </remarks>
        /// <param name="connectionString">�@�ӦX�k��SqlConnection��connection string</param>
        /// <param name="commandType">�R�O�Φ�(stored procedure, text, etc.)</param>
        /// <param name="commandText">stored procedure�W�٩�T-SQL�R�O</param>
        /// <param name="commandParameters">�����檺�R�O��SqlParamters�}�C</param>
        /// <returns>�^�Ǧ��R�O���v�T�����(SqlDataReader)</returns>
        public static SqlDataReader ExecuteReader(string connectionString, CommandType cmdType, string cmdText, params SqlParameter[] commandParameters)
        {
            SqlCommand cmd = new SqlCommand();
            SqlConnection conn = new SqlConnection(connectionString);

            try
            {
                PrepareCommand(cmd, conn, null, cmdType, cmdText, commandParameters);
                SqlDataReader rdr = cmd.ExecuteReader(CommandBehavior.CloseConnection);
                cmd.Parameters.Clear();
                return rdr;
            }
            catch
            {
                conn.Close();
                throw;
            }
        }

        public static XmlReader ExecuteXmlReader(string connectionString, CommandType cmdType, string cmdText, params SqlParameter[] commandParameters)
        {
            SqlCommand cmd = new SqlCommand();
            SqlConnection conn = new SqlConnection(connectionString);

            try
            {
                PrepareCommand(cmd, conn, null, cmdType, cmdText, commandParameters);
                XmlReader rdr = cmd.ExecuteXmlReader();
                cmd.Parameters.Clear();
                return rdr;
            }
            catch
            {
                conn.Close();
                throw;
            }
        }


        /// <summary>
        /// ����@�� SqlCommand (�^�Ǩ��v�T���Ĥ@�Ӹ�ƦC���Ĥ@�����) 
        /// using the provided parameters.
        /// </summary>
        /// <remarks>
        /// e.g.:  
        ///  Object obj = ExecuteScalar(connString, CommandType.StoredProcedure, "PublishOrders", new SqlParameter("@prodid", 24));
        /// </remarks>
        /// <param name="connectionString">�@�ӦX�k��SqlConnection��connection string</param>
        /// <param name="commandType">�R�O�Φ�(stored procedure, text, etc.)</param>
        /// <param name="commandText">stored procedure�W�٩�T-SQL�R�O</param>
        /// <param name="commandParameters">�����檺�R�O��SqlParamters�}�C</param>
        /// <returns>�ϥ�Convert.To{Type}�ӹﵲ�G�૬</returns>
        public static object ExecuteScalar(string connectionString, CommandType cmdType, string cmdText, params SqlParameter[] commandParameters)
        {
            SqlCommand cmd = new SqlCommand();

            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                PrepareCommand(cmd, connection, null, cmdType, cmdText, commandParameters);
                object val = cmd.ExecuteScalar();
                cmd.Parameters.Clear();
                return val;
            }
        }


        /// <summary>
        /// �s�W parameter array �� cache
        /// </summary>
        /// <param name="cacheKey">Key to the parameter cache</param>
        /// <param name="cmdParms">an array of SqlParamters to be cached</param>
        public static void CacheParameters(string cacheKey, params SqlParameter[] commandParameters)
        {
            parmCache[cacheKey] = commandParameters;
        }

        /// <summary>
        /// ���o�bCache����Parameter
        /// </summary>
        /// <param name="cacheKey">key used to lookup parameters</param>
        /// <returns>Cached SqlParamters array</returns>
        public static SqlParameter[] GetCachedParameters(string cacheKey)
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
        /// �ǳư���R�O
        /// </summary>
        /// <param name="cmd">SqlCommand ����</param>
        /// <param name="conn">SqlConnection ����</param>
        /// <param name="trans">SqlTransaction ����</param>
        /// <param name="cmdType">Cmd type e.g. stored procedure or text</param>
        /// <param name="cmdText">Command text, e.g. Select * from Products</param>
        /// <param name="cmdParms">�bcommand�|�Ψ쪺SqlParameters</param>
        private static void PrepareCommand(SqlCommand cmd, SqlConnection conn, SqlTransaction trans, CommandType cmdType, string cmdText, SqlParameter[] cmdParms)
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
                foreach (SqlParameter parm in cmdParms)
                    cmd.Parameters.Add(parm);
            }
        }


    }
}
