using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Data.SqlClient;
using Business_Logic.Common;
using Business_Logic.Entity;
using System.Collections;
using Dapper;
using System.Dynamic;
using System.ComponentModel;
using System.Transactions;
using System.Configuration;

namespace Business_Logic
{
    public class DataComponent
    {
      
        public string strConnectionString = ConfigurationManager.AppSettings["KF_DB"].ToString();



        public SysLog g_SysLog = new SysLog(); 
        public SysEntity.EventLog g_LogEntity = new SysEntity.EventLog();
        public static FunctionHandler g_FunctionHandler = new FunctionHandler();

      

        // 建立帶有參數的建構函式
        public DataComponent(string p_DB ="")
        {
            if (p_DB != "")
            {
                strConnectionString = ConfigurationManager.AppSettings[p_DB].ToString();
            }
        }


        public SysEntity.TransResult ExecuteNonQuery(SysEntity.Employee p_EmployeeEntity, SqlCommand p_cmd, string p_ConnStr = "KF_DB")
        {
            SysEntity.TransResult m_TransResult = new SysEntity.TransResult();
            string m_strConnectionString = ConfigurationManager.AppSettings[p_ConnStr].ToString();
            string m_SQL = p_cmd.CommandText;
            Int32 iExecute = 0;
            try
            {
                using (var conn = new SqlConnection(m_strConnectionString))
                {
                    p_cmd.Connection = conn;
                    p_cmd.Connection.Open();
                    iExecute = p_cmd.ExecuteNonQuery();
                    m_TransResult.isSuccess = true;
                    m_TransResult.ResultEntity = iExecute;

                }
            }
            catch (Exception ex)
            {
                g_LogEntity.LogID = DateTime.Now.ToString("yyyyMMddHHmmssff");
                g_LogEntity.LogResult = "Fail";
                g_LogEntity.LogMessage = ex.StackTrace;
                g_LogEntity.Sql = m_SQL;
                g_LogEntity.Employee = p_EmployeeEntity;
                g_SysLog.EventLogManager(g_LogEntity);
                m_TransResult.isSuccess = false;
                m_TransResult.LogMessage = ex.Message;
                m_TransResult.ResultEntity = ex.StackTrace;
            }

            return m_TransResult;
        }

        public SysEntity.TransResult GetDataTable(SysEntity.Employee p_EmployeeEntity, SqlCommand p_cmd, string p_ConnStr = "KF_DB")
        {

            string m_strConnectionString = ConfigurationManager.AppSettings[p_ConnStr].ToString();
            SysEntity.TransResult m_TransResult = new SysEntity.TransResult();
            DataTable dsDataTable = new DataTable("Table");
            string m_SQL = p_cmd.CommandText;

            try
            {
                using (var conn = new SqlConnection(m_strConnectionString))
                {
                    p_cmd.Connection = conn;
                    using (SqlDataAdapter sda = new SqlDataAdapter(p_cmd))
                    {
                        sda.SelectCommand = p_cmd;
                        sda.Fill(dsDataTable);
                        m_TransResult.isSuccess = true;
                        m_TransResult.ResultEntity = dsDataTable;
                    }
                }
            }
            catch (Exception ex)
            {
                g_LogEntity.LogID = DateTime.Now.ToString("yyyyMMddHHmmssff");
                g_LogEntity.LogResult = "Fail";
                g_LogEntity.LogMessage = ex.StackTrace;
                g_LogEntity.Sql = m_SQL;
                g_LogEntity.Employee = p_EmployeeEntity;
                g_SysLog.EventLogManager(g_LogEntity);
                m_TransResult.isSuccess = false;
                m_TransResult.LogMessage = ex.Message;
            }
            return m_TransResult;
        }

        public SysEntity.TransResult GetData<T>(SysEntity.Employee p_EmployeeEntity, SqlCommand p_cmd, object p_Params, string p_ConnStr = "KF_DB") where T : new()
        {
            SysEntity.TransResult m_TransResult = new SysEntity.TransResult();
            string m_strConnectionString = ConfigurationManager.AppSettings[p_ConnStr].ToString();
            string m_SQL = p_cmd.CommandText;
            try
            {
                using (var conn = new SqlConnection(m_strConnectionString))
                {
                    p_cmd.Connection = conn;
                    m_TransResult.isSuccess = true;
                    m_TransResult.ResultEntity = conn.Query<T>(p_cmd.CommandText, p_Params);
                }
            }
            catch (Exception ex)
            {
                g_LogEntity.LogID = DateTime.Now.ToString("yyyyMMddHHmmssff");
                g_LogEntity.LogResult = "Fail";
                g_LogEntity.LogMessage = ex.StackTrace;
                g_LogEntity.Sql = m_SQL;
                g_LogEntity.Employee = p_EmployeeEntity;
                g_SysLog.EventLogManager(g_LogEntity);
                m_TransResult.isSuccess = false;
                m_TransResult.LogMessage = ex.Message;
            }
            return m_TransResult;
        }

        public DataTable GetDataTable(SqlCommand p_cmd)
        {
            DataTable dsDataTable = new DataTable("Table");
            string m_SQL = p_cmd.CommandText;

            try
            {
                using (var conn = new SqlConnection(strConnectionString))
                {
                    p_cmd.Connection = conn;


                    using (SqlDataAdapter sda = new SqlDataAdapter(p_cmd))
                    {

                        sda.SelectCommand = p_cmd;
                        sda.Fill(dsDataTable);

                    }
                }
            }
            catch (Exception ex)
            {
                g_LogEntity.LogID = DateTime.Now.ToString("yyyyMMddHHmmssff");
                g_LogEntity.LogResult = "Fail";
                g_LogEntity.LogMessage = ex.StackTrace;
                g_LogEntity.Sql = m_SQL;
                g_SysLog.EventLogManager(g_LogEntity);
            }
            return dsDataTable;
        }

        public DataTable GetDataTable(SqlCommand p_cmd, string p_ConnStr = "KF_DB")
        {
          string m_strConnectionString = ConfigurationManager.AppSettings[p_ConnStr].ToString();

            DataTable dsDataTable = new DataTable("Table");
            string m_SQL = p_cmd.CommandText;
            try
            {
                using (var conn = new SqlConnection(m_strConnectionString))
                {
                    p_cmd.Connection = conn;
                    using (SqlDataAdapter sda = new SqlDataAdapter(p_cmd))
                    {
                        sda.SelectCommand = p_cmd;
                        sda.Fill(dsDataTable);
                    }
                }
            }
            catch (Exception ex)
            {
                g_LogEntity.LogID = DateTime.Now.ToString("yyyyMMddHHmmssff");
                g_LogEntity.LogResult = "Fail";
                g_LogEntity.LogMessage = ex.Message;
                g_LogEntity.Sql = m_SQL;
                g_SysLog.EventLogManager(g_LogEntity);
            }
            return dsDataTable;
        }

        public SysEntity.TransResult GridDataCtrl(SysEntity.TransResult p_TransResult, string p_PageSize, Dictionary<string, Object> p_Param)
        {
            string m_GroupBy = "";
            string m_Type = "";
            ArrayList m_alGrid = new ArrayList();
            foreach (KeyValuePair<string, Object> item in p_Param)
            {
                if (item.Key == "GroupBy")
                {
                    m_GroupBy = item.Value.ToString();
                }
                if (item.Key == "Type")
                {
                    m_Type = item.Value.ToString();
                }
            }
            try
            {
                if (m_Type == "ExportExcel")
                {
                    return p_TransResult;
                }
                else
                {
                    if (p_TransResult.isSuccess)
                    {
                        m_alGrid = (ArrayList)p_TransResult.ResultEntity;
                        if (m_alGrid.Count != 0)
                        {
                            m_alGrid[1] = g_FunctionHandler.DataTableToJson((DataTable)m_alGrid[1]);
                            double s = Convert.ToInt32(m_alGrid[0]);
                            int result = 0;
                            result = Convert.ToInt16(Math.Ceiling(s / Convert.ToInt32(p_PageSize)));
                            m_alGrid.Add(result.ToString());
                            m_alGrid.Add(p_PageSize);
                            if (m_GroupBy != "")
                            {
                                m_alGrid.Add(m_GroupBy);
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                p_TransResult.LogMessage = ex.Message;
                p_TransResult.isSuccess = false;
            }
            return p_TransResult;
        }

        public SysEntity.TransResult GetDataTableByGrid(SysEntity.Employee p_EmployeeEntity, SqlCommand p_cmd, Dictionary<string, Object> p_Param, string p_PageSize, string p_PageIndex, string p_ConnStr = "KF_DB")
        {
            string m_OrderBy = "";
            string m_GroupBys = "";
            string m_Type = "";
            SysEntity.TransResult m_TransResult = new SysEntity.TransResult();
            string m_strConnectionString = ConfigurationManager.AppSettings[p_ConnStr].ToString();


            foreach (KeyValuePair<string, Object> item in p_Param)
            {
                if (item.Key == "OrderBy")
                {
                    m_OrderBy = item.Value.ToString();
                }
                if (item.Key == "GroupBy")
                {
                    m_GroupBys = item.Value.ToString();
                }
                if (item.Key == "Type")
                {
                    m_Type = item.Value.ToString();
                }
            }
            if (m_Type == "ExportExcel")
            {

                if (m_OrderBy != "")
                {
                    p_cmd.CommandText = "select * from (" + p_cmd.CommandText + ")M  Order by " + m_OrderBy;
                }

                m_TransResult = GetDataTable(p_EmployeeEntity, p_cmd, p_ConnStr);
            }
            else
            {
                ArrayList m_arl = new ArrayList();
                DataTable dsDataTable = new DataTable("Table");
                string m_SQL = p_cmd.CommandText;

                DataTable m_dtCount = new DataTable("dtCount");
                SqlCommand p_cmdCount = new SqlCommand();

                using (p_cmdCount)
                {
                    using (var conn = new SqlConnection(m_strConnectionString))
                    {
                        p_cmdCount.Connection = conn;
                        using (SqlDataAdapter sda = new SqlDataAdapter(p_cmdCount))
                        {
                            foreach (System.Data.SqlClient.SqlParameter Spc in p_cmd.Parameters)
                            {
                                p_cmdCount.Parameters.Add(new SqlParameter(Spc.ParameterName, Spc.Value));
                            }
                            if (m_GroupBys != "")
                            {
                                p_cmdCount.CommandText = "select count(*)DataCount from (select distinct " + m_GroupBys + " From (" + p_cmd.CommandText + ")Grop ) Data";
                            }
                            else
                            {
                                p_cmdCount.CommandText = "select count(*)DataCount from (" + p_cmd.CommandText + ") Data";
                            }
                            sda.SelectCommand = p_cmdCount;
                            sda.Fill(m_dtCount);
                        }
                    }
                }
                foreach (DataRow dr in m_dtCount.Rows)
                {
                    string DataCount = dr["DataCount"].ToString();
                    m_arl.Add(DataCount);
                }
                try
                {
                    using (var conn = new SqlConnection(m_strConnectionString))
                    {
                        p_cmd.Connection = conn;
                        using (SqlDataAdapter sda = new SqlDataAdapter(p_cmd))
                        {
                            if (m_GroupBys != "")
                            {
                                string m_SqlGroupBys = "";
                                Int32 m_GroupByCount = 0;
                                foreach (string m_GroupBy in m_GroupBys.Split(','))
                                {
                                    if (m_GroupByCount == 0)
                                    {
                                        m_SqlGroupBys += m_GroupBy;
                                    }
                                    else
                                    {
                                        m_SqlGroupBys += "+" + m_GroupBy;
                                    }
                                    m_GroupByCount++;
                                }

                                string m_NewSQL = "";
                                if (p_cmd.CommandText.ToUpper().IndexOf("WHERE") == -1)
                                {
                                    m_NewSQL = p_cmd.CommandText + " where  " + m_SqlGroupBys + "  in  (SELECT  * FROM ( select distinct " + m_SqlGroupBys + " as 'Key' From (" + p_cmd.CommandText + ")Grop) M ORDER BY " + m_SqlGroupBys + " OFFSET (@PageIndex-1)*@PageSize ROWS FETCH NEXT @PageSize ROWS ONLY)  ";
                                }
                                else
                                {
                                    m_NewSQL = p_cmd.CommandText + " and  " + m_SqlGroupBys + "  in  (SELECT  * FROM ( select distinct " + m_SqlGroupBys + " as 'Key' From (" + p_cmd.CommandText + ")Grop) M ORDER BY " + m_SqlGroupBys + " OFFSET (@PageIndex-1)*@PageSize ROWS FETCH NEXT @PageSize ROWS ONLY)  ";
                                }
                                p_cmd.CommandText = m_NewSQL;
                                p_cmd.Parameters.Add(new SqlParameter("PageIndex", Convert.ToInt32(p_PageIndex)));
                                p_cmd.Parameters.Add(new SqlParameter("PageSize", Convert.ToInt32(p_PageSize)));
                            }
                            else
                            {
                                p_cmd.CommandText = " SELECT  * FROM ( " + p_cmd.CommandText + " ) M ORDER BY " + m_OrderBy + " OFFSET (@PageIndex-1)*@PageSize ROWS FETCH NEXT @PageSize ROWS ONLY";
                                p_cmd.Parameters.Add(new SqlParameter("PageIndex", Convert.ToInt32(p_PageIndex)));
                                p_cmd.Parameters.Add(new SqlParameter("PageSize", Convert.ToInt32(p_PageSize)));
                            }

                            if (m_OrderBy != "")
                            {
                                p_cmd.CommandText = "select * from (" + p_cmd.CommandText + ")M  Order by " + m_OrderBy;
                            }
                            sda.SelectCommand = p_cmd;
                            sda.Fill(dsDataTable);
                            m_arl.Add(dsDataTable);
                            m_TransResult.isSuccess = true;
                            m_TransResult.ResultEntity = m_arl;
                        }
                    }
                }
                catch (Exception ex)
                {
                    g_LogEntity.LogID = DateTime.Now.ToString("yyyyMMddHHmmssff");
                    g_LogEntity.LogResult = "Fail";
                    g_LogEntity.LogMessage = ex.StackTrace;
                    g_LogEntity.Sql = m_SQL;
                    g_LogEntity.Employee = p_EmployeeEntity;
                    g_SysLog.EventLogManager(g_LogEntity);
                    m_TransResult.isSuccess = false;
                    m_TransResult.LogMessage = ex.Message;
                }
            }
            return m_TransResult;
        }

        public SysEntity.TransResult GetDataTable(SysEntity.Employee p_EmployeeEntity, SqlCommand p_cmd, string p_OrderBy, string p_PageSize, string p_PageIndex)
        {
            SysEntity.TransResult m_TransResult = new SysEntity.TransResult();
            DataTable dsDataTable = new DataTable("Table");
            string m_SQL = p_cmd.CommandText;

            try
            {
                using (var conn = new SqlConnection(strConnectionString))
                {
                    p_cmd.Connection = conn;


                    using (SqlDataAdapter sda = new SqlDataAdapter(p_cmd))
                    {
                        p_cmd.CommandText = " SELECT  * FROM ( " + p_cmd.CommandText + " ) M ORDER BY [" + p_OrderBy + "] OFFSET (@PageIndex-1)*@PageSize ROWS FETCH NEXT @PageSize ROWS ONLY";
                        p_cmd.Parameters.Add(new SqlParameter("PageIndex", Convert.ToInt32(p_PageIndex)));
                        p_cmd.Parameters.Add(new SqlParameter("PageSize", Convert.ToInt32(p_PageSize)));


                        sda.SelectCommand = p_cmd;
                        sda.Fill(dsDataTable);
                        m_TransResult.isSuccess = true;
                        m_TransResult.ResultEntity = dsDataTable;

                    }
                }
            }
            catch (Exception ex)
            {
                g_LogEntity.LogID = DateTime.Now.ToString("yyyyMMddHHmmssff");
                g_LogEntity.LogResult = "Fail";
                g_LogEntity.LogMessage = ex.StackTrace;
                g_LogEntity.Sql = m_SQL;
                g_LogEntity.Employee = p_EmployeeEntity;
                g_SysLog.EventLogManager(g_LogEntity);
                m_TransResult.isSuccess = false;
                m_TransResult.LogMessage = ex.Message;
            }
            return m_TransResult;
        }

        public DataSet GetDataSet(SysEntity.Employee p_EmployeeEntity, SqlCommand p_cmd)
        {
            string m_SQL = p_cmd.CommandText;

            DataSet dsDataSet = new DataSet("DataSet");
            try
            {
                using (var conn = new SqlConnection(strConnectionString))
                {
                    conn.Open();
                    using (SqlDataAdapter sda = new SqlDataAdapter())
                    {
                        using (p_cmd)
                        {
                            sda.SelectCommand = p_cmd;
                            sda.Fill(dsDataSet);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                g_LogEntity.LogID = DateTime.Now.ToString("yyyyMMddHHmmssff");
                g_LogEntity.LogResult = "Fail";
                g_LogEntity.LogMessage = ex.StackTrace;
                g_LogEntity.Sql = m_SQL;
                g_LogEntity.Employee = p_EmployeeEntity;
                g_SysLog.EventLogManager(g_LogEntity);
            }
            return dsDataSet;
        }

        /// <summary>
        /// 共用查詢ByEntity
        /// </summary>
        /// <param name="p_SQL">SQL語法</param>
        /// <param name="p_Entity">Entity</param>
        /// <returns></returns>
        public SqlCommand GetQueryCommand(SysEntity.Employee p_EmployeeEntity, string p_SQL, object p_Entity, Dictionary<string, Object> p_BetWeen = null, string[] p_LikeProp = null, string p_OrderBy = "")
        {
            string m_Where = "";
            SqlCommand m_cmd = new SqlCommand("SELECT * FROM (" + p_SQL + " ) Master ");
            
            bool m_isDictionary = false;

            ArrayList m_Entity = new ArrayList();


            if (p_Entity.GetType().Name.IndexOf("Dictionary") != 0)
            {
                m_Entity = new ArrayList(p_Entity.GetType().GetProperties());
            }
            else
            {
                m_isDictionary = true;
                m_Entity = new ArrayList((Dictionary<string, Object>)p_Entity);
            }

            foreach (var prop in m_Entity)
            {
                string m_propName = "";
                object value = new object();
                if (p_Entity.GetType().Name.IndexOf("Dictionary") != 0)
                {
                    m_propName = ((System.Reflection.PropertyInfo)(prop)).Name;
                    value = ((System.Reflection.PropertyInfo)(prop)).GetValue(p_Entity, null);
                }
                else
                {
                    m_propName = ((KeyValuePair<string, Object>)prop).Key;
                    if (((KeyValuePair<string, Object>)prop).Value != null)
                    {
                        value = ((KeyValuePair<string, Object>)prop).Value.ToString();
                    }
                    else
                    {
                        value = "";
                    }
                }
                var m_SiteFormStatusConditionType = "";
                if (m_propName == "CommonSiteFormStatus")
                {
                    string m_CommonSiteFormStatus = value.ToString();
                    if (m_CommonSiteFormStatus != "")
                    {
                        if (m_CommonSiteFormStatus.Split(';').Length != 0)
                        {
                            m_propName = m_CommonSiteFormStatus.Split(';')[0];
                            if (m_propName.IndexOf("||") == -1)
                            {
                                m_SiteFormStatusConditionType = m_CommonSiteFormStatus.Split(';')[1];
                                value = m_CommonSiteFormStatus.Split(';')[2];
                                if (value.ToString() == "LoginUser")
                                {
                                    value = p_EmployeeEntity.WorkID;
                                }
                            }
                            else
                            {
                            }
                        }
                    }
                }
                bool m_isBetWeen = false;
                bool m_isBetWeenEnd = false;
                if (m_isDictionary)
                {
                    if (m_propName.IndexOf("_S") == m_propName.Length - 2)
                    {
                        m_isBetWeen = true;
                    }
                    if (m_propName.IndexOf("_E") == m_propName.Length - 2)
                    {
                        m_isBetWeen = true;
                        m_isBetWeenEnd = true;

                    }
                }
                else
                {
                    m_isBetWeen = CheckBetWeen(p_BetWeen, m_propName);
                }

                if (m_isBetWeen)
                {
                    if (!m_isBetWeenEnd)
                    {
                        string m_BetWeen_S = "";
                        string m_BetWeen_E = "";
                        if (m_isDictionary)
                        {
                            string m_BetWeenEnd_propName = m_propName.Remove(m_propName.IndexOf("_S")) + "_E";
                            m_BetWeen_S = p_BetWeen[m_propName].ToString();
                            m_BetWeen_E = p_BetWeen[m_BetWeenEnd_propName].ToString();
                            if (m_BetWeen_E != "" && m_BetWeen_E != "")
                            {
                                if (m_Where == "")
                                {
                                    m_Where += " WHERE " + m_propName.Remove(m_propName.IndexOf("_S")) + " BetWeen @" + m_propName + " AND @" + m_BetWeenEnd_propName;
                                }
                                else
                                {
                                    m_Where += " AND " + m_propName.Remove(m_propName.IndexOf("_S")) + " BetWeen @" + m_propName + " AND @" + m_BetWeenEnd_propName;
                                }
                                m_cmd.Parameters.Add(new SqlParameter(m_propName, m_BetWeen_S));
                                m_cmd.Parameters.Add(new SqlParameter(m_BetWeenEnd_propName, m_BetWeen_E));
                            }
                        }
                        else
                        {
                            m_BetWeen_S = p_BetWeen[m_propName + "_S"].ToString();
                            m_BetWeen_E = p_BetWeen[m_propName + "_E"].ToString();
                            if (m_BetWeen_S != "" && m_BetWeen_E != "")
                            {
                                if (m_Where == "")
                                {
                                    m_Where += " WHERE " + m_propName + " BetWeen @" + m_propName + "_S AND @" + m_propName + "_E ";
                                }
                                else
                                {
                                    m_Where += " AND " + m_propName + " BetWeen @" + m_propName + "_S AND @" + m_propName + "_E ";
                                }
                            }
                            m_cmd.Parameters.Add(new SqlParameter(m_propName + "_S", m_BetWeen_S));
                            m_cmd.Parameters.Add(new SqlParameter(m_propName + "_E", m_BetWeen_E));
                        }
                    }
                }
                else
                {
                    string m_Equation = "=@" + m_propName;
                    if (m_propName.IndexOf("NOTIN_") != -1)
                    {
                        m_propName = m_propName.Replace("NOTIN_", "");
                        m_Equation = " <> @" + m_propName;
                    }


                    bool m_isLike = false;
                    if (p_LikeProp != null)
                    {
                        if (p_LikeProp.Contains(m_propName))
                        {
                            m_Equation = " like @" + m_propName;
                            m_isLike = true;
                        }
                    }


                    if (value != null)
                    {
                        if (value.ToString().ToLower().IndexOf("between") == 0)
                        {
                        }
                        string m_EquationVal = value.ToString();

                        if (m_EquationVal != "")
                        {
                            if (m_isLike)
                            {
                                m_EquationVal = "%" + value.ToString() + "%";
                            }

                            string m_WhereValue = "";
                            if (m_SiteFormStatusConditionType != "between")
                            {
                                m_WhereValue = m_propName + m_Equation;
                            }
                            else
                            {
                                m_WhereValue = m_propName + "  " + m_SiteFormStatusConditionType + " '" + m_EquationVal.Split(',')[0] + "' and '" + m_EquationVal.Split(',')[1] + "'";
                            }
                            if (m_Where == "")
                            {
                                m_Where += " WHERE " + m_WhereValue;
                            }
                            else
                            {
                                m_Where += " AND " + m_WhereValue;
                            }
                            if (m_SiteFormStatusConditionType != "between")
                            {
                                m_cmd.Parameters.Add(new SqlParameter(m_propName, m_EquationVal));
                            }
                        }
                    }
                }
            }
            m_cmd.CommandText += m_Where +" "+ p_OrderBy;

            return m_cmd;
        }


        public bool CheckBetWeen(Dictionary<string, Object> p_BetWeen, string p_propName)
        {
            try
            {
                if (p_BetWeen!=null)
                {
                    p_BetWeen[p_propName + "_S"].ToString();
                    return true;
                }
                else
                {
                    return false;
                }
               
               
            }
            catch
            {
                return false;
            }
        }

        public SysEntity.TransResult TransInsertSQL(SysEntity.Employee p_Employee, string p_TableName, DataTable p_dt, Dictionary<string, string> p_Params = null)
        {
            SysEntity.TransResult m_TransResult = new SysEntity.TransResult();
            try
            {
                if (p_dt.Rows.Count != 0)
                {
                    ArrayList m_arrSQL = new ArrayList();
                    Dictionary<string, object> m_dicSql = new Dictionary<string, object>();

                    List<SqlParameter> m_Parameters = new List<SqlParameter>();
                    string m_DelSql = " Delete FROM [dbo].[" + p_TableName + "] ";
                    string m_DelSqlWhere = "";
                    Int32 m_ParamCount = 0;
                    if (p_Params != null)
                    {
                        foreach (KeyValuePair<string, string> item in p_Params)
                        {
                            if (m_DelSqlWhere == "")
                            {
                                m_DelSqlWhere += " Where " + item.Key + " = @" + item.Key + m_ParamCount.ToString();
                            }
                            else
                            {
                                m_DelSqlWhere += " And " + item.Key + "= @" + item.Key + m_ParamCount.ToString();

                            }
                            m_Parameters.Add(new SqlParameter(item.Key + m_ParamCount.ToString(), item.Value.Replace("''", "'")));
                            m_ParamCount++;
                        }
                    }

                    m_dicSql["SQL"] = m_DelSql + m_DelSqlWhere;
                    m_dicSql["Params"] = m_Parameters;
                    m_arrSQL.Add(m_dicSql);

                    Int32 m_CommitCount = 0;
                    using (SqlConnection conn = new SqlConnection(strConnectionString))
                    {
                        conn.Open();
                        using (SqlCommand command = new SqlCommand())
                        {
                            SqlTransaction transaction;
                            //Start Transaction
                            transaction = conn.BeginTransaction();
                            command.Connection = conn;
                            command.Transaction = transaction;
                            try
                            {
                                foreach (Dictionary<string, object> Sql in m_arrSQL)
                                {
                                    command.CommandText = Sql["SQL"].ToString();
                                    foreach (SqlParameter Param in (List<SqlParameter>)Sql["Params"])
                                    {
                                        command.Parameters.Add(Param);
                                    }
                                    m_CommitCount += command.ExecuteNonQuery();
                                }
                                using (SqlBulkCopy bulkCopy = new SqlBulkCopy(conn, SqlBulkCopyOptions.KeepIdentity, transaction))
                                {
                                    //DataTable dt = GetTempTableData();
                                    bulkCopy.DestinationTableName = "dbo." + p_TableName;
                                    foreach (DataColumn dc in p_dt.Columns)
                                    {
                                        bulkCopy.ColumnMappings.Add(dc.ColumnName, dc.ColumnName);
                                    }
                                    try
                                    {
                                        bulkCopy.WriteToServer(p_dt);
                                        m_TransResult.isSuccess = true;

                                    }
                                    catch (Exception ex)
                                    {
                                        m_TransResult.isSuccess = false;
                                        m_TransResult.LogMessage = ex.Message;
                                        transaction.Rollback();
                                    }
                                }
                                if (m_TransResult.isSuccess)
                                {
                                    transaction.Commit();
                                    m_TransResult.isSuccess = true;
                                    m_TransResult.ResultEntity = m_CommitCount;
                                }
                            }
                            catch (Exception ex)
                            {
                                m_TransResult.isSuccess = false;
                                m_TransResult.LogMessage = ex.Message;
                                //有問題即Rollback
                                transaction.Rollback();
                            }
                            command.Connection.Close();
                        }

                    }
                }
                

            }
            catch
            {
                throw;
            }
            return m_TransResult;
        }


        public SysEntity.TransResult InsertIntoByEntity<T>(SysEntity.Employee p_EmployeeEntity, string p_TableName, Dictionary<string, Object> p_Entity, Boolean p_isCreateUser=true, string p_ConnStr = "KF_DB") where T : new()
        {
            var m_STypet = new T();
            SysEntity.TransResult m_TransResult = new SysEntity.TransResult();
            try
            {
                var props = m_STypet.GetType().GetProperties();
                string m_SQL = "Insert into [dbo].[" + p_TableName + "]  ";

                string m_Columns = "";
                string m_Values = "";
                object m_Obj = null;
                using (SqlCommand m_cmd = new SqlCommand())
                {
                    foreach (var prop in props)
                    {
                        string m_propName = prop.Name;
                        if (p_Entity.TryGetValue(m_propName, out m_Obj))
                        {
                            if (p_Entity[m_propName] != null)
                            {
                                string m_value = p_Entity[m_propName].ToString(); // against prop.Name
                                if (m_value != "")
                                {
                                    m_Columns += (m_Columns == "" ? "(" : ",") + m_propName;
                                    m_Values += (m_Values == "" ? "Values(" : ",") + "@" + m_propName;
                                    m_cmd.Parameters.Add(new SqlParameter(m_propName, m_value.ToString()));
                                }
                            }
                        }
                    }
                    if (p_isCreateUser)
                    {
                        if (p_ConnStr == "KF_DB")
                        {
                            m_SQL += m_Columns + ",CreateUser,ModifyUser )  " + m_Values + ",'" + p_EmployeeEntity.WorkID + "','" + p_EmployeeEntity.WorkID + "' ) ";

                        }
                        else
                        { 
                        m_SQL += m_Columns + ",Add_User,Upd_User )  " + m_Values + ",'" + p_EmployeeEntity.WorkID + "','" + p_EmployeeEntity.WorkID + "' ) ";

                        }
                    }
                    else
                    {
                        m_SQL += m_Columns + " )  " + m_Values + " ) ";
                    
                    }
                    m_cmd.CommandText = m_SQL;
                    m_TransResult = ExecuteNonQuery(p_EmployeeEntity, m_cmd, p_ConnStr);
                }
            }
            catch (Exception ex)
            {
                m_TransResult.LogMessage = ex.Message;
                m_TransResult.isSuccess = false;
            }

            return m_TransResult;
        }

        public dynamic ConvertToDynamic(Dictionary<string, Object> obj)
        {
            IDictionary<string, object> result = new ExpandoObject();

            foreach (KeyValuePair<string, Object> item in obj)
            {
                result.Add(item.Key, item.Value);
            }

            return result as ExpandoObject;
        }

        public SysEntity.CommandEntity GetInsertCommand(SysEntity.Employee p_EmployeeEntity, string p_TableName, Dictionary<string, Object> p_Entity)
        {
            SysEntity.CommandEntity m_CommandEntity = new SysEntity.CommandEntity();
            try
            {
                string m_SQL = "Insert into [dbo].[" + p_TableName + "]  ";
                List<object> m_lisParams = new List<object>();

                string m_Columns = "";
                string m_Values = "";
                object m_Obj = null;

                object m_Param = new object();
                
                ArrayList m_Entity = new ArrayList((Dictionary<string, Object>)p_Entity);
                foreach (var prop in m_Entity)
                {
                    string m_propName = "";
                    object value = new object();
                    if (p_Entity.GetType().Name.IndexOf("Dictionary") != 0)
                    {
                        m_propName = ((System.Reflection.PropertyInfo)(prop)).Name;
                        value = ((System.Reflection.PropertyInfo)(prop)).GetValue(p_Entity, null);
                    }
                    else
                    {
                        m_propName = ((KeyValuePair<string, Object>)prop).Key;
                        value = ((KeyValuePair<string, Object>)prop).Value.ToString();
                    }

                    if (p_Entity.TryGetValue(m_propName, out m_Obj))
                    {
                        string m_value = p_Entity[m_propName].ToString(); // against prop.Name
                        if (m_value != "")
                        {
                            m_Columns += (m_Columns == "" ? "(" : ",") + m_propName;
                            m_Values += (m_Values == "" ? "Values(" : ",") + "@" + m_propName;
                        }
                    }
                }

                object exProd = ConvertToDynamic(p_Entity);
                m_CommandEntity.Parameters = exProd;
                m_SQL += m_Columns + " )  " + m_Values + " ) ";
                m_CommandEntity.SQLCommand = m_SQL;

            }
            catch (Exception ex)
            {
                throw ex;
            }
            return m_CommandEntity;
        }

        public SysEntity.TransResult ExecuteCommandEntitys(SysEntity.Employee p_Employee, List<SysEntity.CommandEntity> p_CommandEntitys)
        {
            SysEntity.TransResult m_TransResult = new SysEntity.TransResult();
            Int32 m_CommitCount = 0;
            using (var transactionScope = new TransactionScope())
            {
                using (SqlConnection conn = new SqlConnection(strConnectionString))
                {
                    //Start Transaction
                    conn.Open();
                    try
                    {
                        foreach (SysEntity.CommandEntity CommandEntity in p_CommandEntitys)
                        {
                            m_CommitCount += conn.Execute(CommandEntity.SQLCommand, CommandEntity.Parameters);
                        }
                        transactionScope.Complete();
                        m_TransResult.isSuccess = true;
                        m_TransResult.ResultEntity = m_CommitCount;

                    }
                    catch (Exception ex)
                    {
                        m_TransResult.isSuccess = false;
                        m_TransResult.LogMessage = ex.Message;
                    }
                }
            }

            return m_TransResult;
        }

        public SysEntity.TransResult QueryCommandEntitys<T>(SysEntity.Employee p_Employee, SysEntity.CommandEntity p_CommandEntity) where T : new()
        {
            var m_STypet = new T();
            SysEntity.TransResult m_TransResult = new SysEntity.TransResult();
            using (SqlConnection conn = new SqlConnection(strConnectionString))
            {
                try
                {
                    m_TransResult.ResultEntity = conn.Query<T>(p_CommandEntity.SQLCommand, p_CommandEntity.Parameters);
                    m_TransResult.isSuccess = true;


                }
                catch (Exception ex)
                {
                    m_TransResult.isSuccess = false;
                    m_TransResult.LogMessage = ex.Message;
                }
            }


            return m_TransResult;
        }


    }
}
