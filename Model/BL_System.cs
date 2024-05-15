using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Business_Logic.Entity;
using System.Data.SqlClient;
using System.Data;
using System.Web.UI.WebControls;
using System.Collections;
using Newtonsoft.Json;
using System.Reflection;
using System.Xml;
using System.Configuration;
using static Business_Logic.Entity.SysEntity;
using NPOI.SS.Formula.Functions;
using System.Security.Policy;
using System.Text.RegularExpressions;


namespace Business_Logic.Model
{
    public class BL_System
    {
        DataComponent g_DC = new DataComponent();

        public static FunctionHandler g_FH = new FunctionHandler();

        public SysEntity.TransResult ChkIP(string IP)
        {
            SysEntity.TransResult m_TransResult = new SysEntity.TransResult();
            m_TransResult.isSuccess = true;
            try
            {
                SysEntity.Employee m_EmployeeEntity = new SysEntity.Employee();
                m_EmployeeEntity.WorkID = "ChkIP";  

                string m_SQL = " select count(*) LoginCount from Login_Log where RequestIP=@IP " +
                               "  and RequestTime between DATEADD(MI,-3,GETDATE() ) and  GETDATE() and IsPass='N' " +
                               " having count(*) >= 5 ";
                using (SqlCommand m_cmd = new SqlCommand(m_SQL))
                {
                    m_cmd.Parameters.Add(new SqlParameter("IP", IP));

                    m_TransResult = g_DC.GetDataTable(m_EmployeeEntity, m_cmd);
                    if (m_TransResult.isSuccess)
                    {
                        if (((DataTable)m_TransResult.ResultEntity).Rows.Count!=0)
                        {
                            m_TransResult.isSuccess = false;
                            m_TransResult.LogMessage = "錯誤太多次,請稍後再試";
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                m_TransResult.LogMessage = ex.Message;
                m_TransResult.isSuccess = false;
            }
            return m_TransResult;
        }
        /// <summary>
        /// TimeOut重新登入
        /// </summary>
        /// <param name="WorkID"></param>
        /// <param name="PW"></param>
        /// <returns></returns>
        public SysEntity.TransResult ReLogin( string WorkID, string PW, string Company)
        {
            SysEntity.TransResult m_TransResult = new SysEntity.TransResult();
            try
            {
                //fffff
                SysEntity.Employee m_EmployeeEntity = new SysEntity.Employee();
                m_EmployeeEntity.WorkID = "ReLogin";
                if (Company != "AE")
                {
                    string m_SQL = " select E.*,G.GroupID from EmployeeInfo E Left Join EmployeeGroup G on E.WorkID=G.WorkID where E.WorkID=@WorkID  and Password=@PW  ";
                    using (SqlCommand m_cmd = new SqlCommand(m_SQL))
                    {
                        m_cmd.Parameters.Add(new SqlParameter("WorkID", WorkID));
                        m_cmd.Parameters.Add(new SqlParameter("PW", PW));
                        m_TransResult = g_DC.GetDataTable(m_EmployeeEntity, m_cmd);
                        if (m_TransResult.isSuccess)
                        {
                            DataTable dtEmployee = (DataTable)m_TransResult.ResultEntity;
                            if (dtEmployee.Rows.Count > 0)
                            {
                                SysEntity.Employee m_ResultEmployee = new SysEntity.Employee();
                                foreach (DataRow dr in dtEmployee.Rows)
                                {
                                    m_ResultEmployee.CurrentLanguage = "TW";
                                    m_ResultEmployee.WorkID = WorkID;
                                    m_ResultEmployee.DisplayName = dr["DisplayName"].ToString();
                                    m_ResultEmployee.CompanyCode = dr["CompanyCode"].ToString();
                                    m_ResultEmployee.Remark = dr["Remark"].ToString();
                                    m_ResultEmployee.GroupID = dr["GroupID"].ToString();
                                }
                                m_TransResult.ResultEntity = (SysEntity.Employee)m_ResultEmployee;

                            }
                            else
                            {
                                m_TransResult.isSuccess = false;
                                m_TransResult.LogMessage = "輸入帳號密碼錯誤!!";
                            }
                        }
                    }
                }
                else
                {
                    string m_SQL = " select * from User_M E  where E.U_num=@WorkID and U_psw=@PW  ";
                    using (SqlCommand m_cmd = new SqlCommand(m_SQL))
                    {
                        m_cmd.Parameters.Add(new SqlParameter("WorkID", WorkID));
                        m_cmd.Parameters.Add(new SqlParameter("PW", PW));
                        m_TransResult = g_DC.GetDataTable(m_EmployeeEntity, m_cmd, "AE_DB");
                        if (m_TransResult.isSuccess)
                        {
                            DataTable dtEmployee = (DataTable)m_TransResult.ResultEntity;
                            if (dtEmployee.Rows.Count > 0)
                            {
                                SysEntity.Employee m_ResultEmployee = new SysEntity.Employee();
                                foreach (DataRow dr in dtEmployee.Rows)
                                {
                                    m_ResultEmployee.CurrentLanguage = "TW";
                                    m_ResultEmployee.WorkID = WorkID;
                                    m_ResultEmployee.DisplayName = dr["U_name"].ToString();
                                    m_ResultEmployee.CompanyCode = "AE";
                                    m_ResultEmployee.GroupID = "AE_DB";
                                    m_ResultEmployee.Remark = "AE";
                                }
                                m_TransResult.ResultEntity = (SysEntity.Employee)m_ResultEmployee;
                            }
                            else
                            {
                                m_TransResult.isSuccess = false;
                                m_TransResult.LogMessage = "輸入帳號密碼錯誤!!";
                            }
                        }
                    }
                }
                

            }
            catch (Exception ex)
            {
                m_TransResult.LogMessage = ex.Message;
                m_TransResult.isSuccess = false;
            }
            return m_TransResult;
        }

        public SysEntity.TransResult ChkLogin(string IP, string WorkID,string PW)
        {
            SysEntity.TransResult m_TransResult = new SysEntity.TransResult();
            try
            {
                SysEntity.Employee m_EmployeeEntity = new SysEntity.Employee ();
                m_EmployeeEntity.WorkID = "LoginChk";
                //檢查IP是否短時間登入多次
                m_TransResult = ChkIP(IP);
                if (m_TransResult.isSuccess)
                {
                    //刪除30天前的登入資料
                    string m_DelSQL = "Delete [dbo].[Login_Log] Where  RequestTime < DATEADD(Day,-30,GETDATE())  ";
                    using (SqlCommand m_Delcmd = new SqlCommand(m_DelSQL))
                    {
                        g_DC.ExecuteNonQuery(m_EmployeeEntity, m_Delcmd);
                    }


                    //登入寫入Log
                    Dictionary<string, Object> m_Login_Log = new Dictionary<string, object>();
                    m_Login_Log.Add("RequestIP", IP);
                    m_Login_Log.Add("WorkID", WorkID);
                   
                    string m_SQL = " select E.[WorkID],E.[DisplayName],CompanyCode,Remark,G.GroupID from EmployeeInfo E Left Join EmployeeGroup G on E.WorkID=G.WorkID where E.WorkID=@WorkID and Password=@PW  ";
                    using (SqlCommand m_cmd = new SqlCommand(m_SQL))
                    {
                        m_cmd.Parameters.Add(new SqlParameter("WorkID", WorkID));
                        m_cmd.Parameters.Add(new SqlParameter("PW", PW));
                        m_TransResult = g_DC.GetDataTable(m_EmployeeEntity, m_cmd);
                        if (m_TransResult.isSuccess)
                        {
                            DataTable dtEmployee = (DataTable)m_TransResult.ResultEntity;
                            if (dtEmployee.Rows.Count > 0)
                            {
                                m_Login_Log.Add("IsPass", "Y");
                                SysEntity.Employee m_ResultEmployee = new SysEntity.Employee();
                                foreach (DataRow dr in dtEmployee.Rows)
                                {
                                    m_ResultEmployee.CurrentLanguage = "TW";
                                    m_ResultEmployee.WorkID = WorkID;
                                    m_ResultEmployee.DisplayName = dr["DisplayName"].ToString();
                                    m_ResultEmployee.CompanyCode = dr["CompanyCode"].ToString();
                                    m_ResultEmployee.GroupID = dr["GroupID"].ToString();
                                    m_ResultEmployee.Remark = dr["Remark"].ToString();
                                }
                                m_TransResult.ResultEntity = (SysEntity.Employee)m_ResultEmployee;
                               
                            }
                            else
                            {
                                m_Login_Log.Add("IsPass", "N");
                                m_TransResult.isSuccess = false;
                                m_TransResult.LogMessage = "輸入帳號密碼錯誤!!";
                            }
                            g_DC.InsertIntoByEntity<SysEntity.Login_Log>(m_EmployeeEntity, "Login_Log", m_Login_Log, false);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                m_TransResult.LogMessage = ex.Message;
                m_TransResult.isSuccess = false;
            }
            return m_TransResult;
        }

        public SysEntity.TransResult SaveEmployee(SysEntity.Employee p_EmployeeEntity,  string Password)
        {
            SysEntity.TransResult m_TransResult = new SysEntity.TransResult();
            try
            {
                string m_MSG = "";
                string m_UpdColumn = " set DisplayName=@DisplayName,Remark=@Remark  ";
                if (Password != "")
                {
                    m_UpdColumn += ",Password=@Password";
                }
                string m_SQL = "update  [dbo].[EmployeeInfo]  "+ m_UpdColumn  + " Where  WorkID= @WorkID  ";
                using (SqlCommand m_cmd = new SqlCommand(m_SQL))
                {
                    m_cmd.Parameters.Add(new SqlParameter("WorkID", p_EmployeeEntity.WorkID));
                    m_cmd.Parameters.Add(new SqlParameter("DisplayName", p_EmployeeEntity.DisplayName));
                    m_cmd.Parameters.Add(new SqlParameter("Remark", p_EmployeeEntity.Remark));

                    if (Password != "")
                    {
                        m_cmd.Parameters.Add(new SqlParameter("Password", Password));
                        m_MSG = "密碼更新成功";
                    }

                    m_TransResult = g_DC.ExecuteNonQuery(p_EmployeeEntity, m_cmd);
                    if (m_TransResult.isSuccess)
                    {
                        if (m_MSG != "")
                        {
                            m_TransResult.LogMessage = m_MSG;

                        }
                        else
                        {
                            m_TransResult.LogMessage = "更新成功";
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                m_TransResult.LogMessage = ex.Message;
                m_TransResult.isSuccess = false;
            }
            return m_TransResult;
        }

        public SysEntity.TransResult GetEventSetting(SysEntity.Employee p_EmployeeEntity, SysEntity.EventSetting p_EventSetting)
        {
            SysEntity.TransResult m_TransResult = new SysEntity.TransResult();
            try
            {
                string m_SQL = "SELECT [EventID],[EventModel],[EventAction],[EventDescription],[EventActionType],[EventRef],[EventCtrl] FROM [dbo].[EventSetting]   ";
                using (SqlCommand m_cmd = g_DC.GetQueryCommand(p_EmployeeEntity, m_SQL, p_EventSetting))
                {
                    m_TransResult = g_DC.GetDataTable(p_EmployeeEntity, m_cmd);
                }
            }
            catch (Exception ex)
            {
                m_TransResult.LogMessage = ex.Message;
                m_TransResult.isSuccess = false;
            }
            return m_TransResult;
        }

        public SysEntity.TransResult GetSelSouceByEventIDs(SysEntity.Employee p_EmployeeEntity, List<Dictionary<string, Object>> p_EventIDs)
        {
            SysEntity.TransResult m_TransResult = new SysEntity.TransResult();
            List<Dictionary<string, Object>> disResults = new List<Dictionary<string, Object>>();
            Dictionary<string, Object> disResult = new Dictionary<string, Object>();

            try
            {
                foreach (Dictionary<string, Object> m_EventID in p_EventIDs)
                {
                    SysEntity.EventSetting m_EventEntry = new SysEntity.EventSetting();
                    m_EventEntry.EventID = m_EventID["selQID"].ToString();
                    m_TransResult = GetEventSetting(p_EmployeeEntity, m_EventEntry);
                    string m_EventModel = "";
                    string m_EventAction = "";
                    string m_EventActionType = "";
                    if (m_TransResult.isSuccess)
                    {
                        foreach (DataRow dr in ((DataTable)m_TransResult.ResultEntity).Rows)
                        {
                            m_EventModel = dr["EventModel"].ToString();
                            m_EventAction = dr["EventAction"].ToString();
                            m_EventActionType = dr["EventActionType"].ToString();
                        }
                    }
                    Type m_Type = g_FH.GetBLType(m_EventModel);

                    object instance = Activator.CreateInstance(m_Type);
                    MethodInfo method = instance.GetType().GetMethod(m_EventAction);
                    object[] m_Params = new object[2];
                    m_Params[0] = p_EmployeeEntity;
                    m_Params[1] = null;
                    m_TransResult = (SysEntity.TransResult)method.Invoke(instance, m_Params);
                    if (m_TransResult.isSuccess)
                    {
                        DataTable m_dtEvent = (DataTable)m_TransResult.ResultEntity;
                        disResult = new Dictionary<string, Object>();
                        disResult["id"] = m_EventID["id"].ToString();
                        ArrayList m_arrText = new ArrayList();
                        foreach (DataRow dr in m_dtEvent.Rows)
                        {
                            m_arrText.Add(dr[m_EventID["selText"].ToString()].ToString());
                        }
                        disResult["Result"] = m_arrText;
                        disResults.Add(disResult);
                    }
                }
                m_TransResult.isSuccess = true;
                m_TransResult.ResultEntity = disResults;


            }
            catch (Exception ex)
            {
                m_TransResult.isSuccess = false;
                m_TransResult.LogMessage = ex.Message;
            }
            return m_TransResult;

        }

        public SysEntity.TransResult GetAreaCode(SysEntity.Employee p_EmployeeEntity)
        {
            SysEntity.TransResult m_TransResult = new SysEntity.TransResult();
            try
            {
                string m_SQL = "SELECT  item_D_code, item_D_name, item_sort FROM  Item_list WHERE  item_M_code = 'id_areacode' AND item_D_type = 'Y' ORDER BY   item_sort   ";
                using (SqlCommand m_cmd = new SqlCommand(m_SQL))
                {

                    DataTable m_DataTable = g_DC.GetDataTable(m_cmd, "AE_DB");
                    m_TransResult.isSuccess = true;
                    m_TransResult.ResultEntity =(m_DataTable);      
                }
            }
            catch (Exception ex)
            {
                m_TransResult.LogMessage = ex.Message;
                m_TransResult.isSuccess = false;
            }
            return m_TransResult;
        }


        public SysEntity.TransResult GetEmployeeGrid(SysEntity.Employee p_EmployeeEntity, Dictionary<string, Object> p_dicEmployee, string p_PageIndex, string p_PageSize, Dictionary<string, Object> p_Param)
        {
            SysEntity.TransResult m_TransResult = new SysEntity.TransResult();
            Dictionary<string, Object> m_objBetween = new Dictionary<string, Object>();
          
          
            ArrayList m_alGrid = new ArrayList();

          
            string[] m_LikeProp = { "WorkID", "DisplayName"};

            try
            {
                string m_SQL = " SELECT WorkID,DisplayName   FROM dbo.EmployeeInfo  ";
                using (SqlCommand m_cmd = g_DC.GetQueryCommand(p_EmployeeEntity, m_SQL, p_dicEmployee, m_objBetween, m_LikeProp))
                {
                    m_TransResult = g_DC.GetDataTableByGrid(p_EmployeeEntity, m_cmd, p_Param, p_PageSize, p_PageIndex);
                    if (m_TransResult.isSuccess)
                    {
                        m_TransResult = g_DC.GridDataCtrl(m_TransResult, p_PageSize, p_Param);
                    }
                }
            }
            catch (Exception ex)
            {
                m_TransResult.LogMessage = ex.Message;
                m_TransResult.isSuccess = false;
            }

            return m_TransResult;
        }

        public SysEntity.TransResult GetAPILogGrid(SysEntity.Employee p_EmployeeEntity, Dictionary<string, Object> p_dicEmployee, string p_PageIndex, string p_PageSize, Dictionary<string, Object> p_Param)
        {
            SysEntity.TransResult m_TransResult = new SysEntity.TransResult();
            Dictionary<string, Object> m_objBetween = new Dictionary<string, Object>();


            ArrayList m_alGrid = new ArrayList();


            string[] m_LikeProp = {  };

            try
            {
                string m_SQL = "   select [TransactionId],[API_Name],[Form_No],[ResultJSON],[CallUser],[CallTime],[StatusCode] from [tbAPILog]  ";
                using (SqlCommand m_cmd = g_DC.GetQueryCommand(p_EmployeeEntity, m_SQL, p_dicEmployee, m_objBetween, m_LikeProp))
                {
                    m_TransResult = g_DC.GetDataTableByGrid(p_EmployeeEntity, m_cmd, p_Param, p_PageSize, p_PageIndex);
                    if (m_TransResult.isSuccess)
                    {
                        m_TransResult = g_DC.GridDataCtrl(m_TransResult, p_PageSize, p_Param);
                    }
                }
            }
            catch (Exception ex)
            {
                m_TransResult.LogMessage = ex.Message;
                m_TransResult.isSuccess = false;
            }

            return m_TransResult;
        }


        #region EmployeeInfo Methods

        public SysEntity.TransResult GetEmployeeInfos(SysEntity.Employee p_EmployeeEntity, Dictionary<string, Object> p_dicEmployeeInfo)
        {
            SysEntity.TransResult m_TransResult = new SysEntity.TransResult();
            Dictionary<string, Object> m_objBetween = new Dictionary<string, Object>();
            SysEntity.EmployeeInfo m_EmployeeInfo = g_FH.ConvertEntity<SysEntity.EmployeeInfo>(p_dicEmployeeInfo, ref m_objBetween);

            try
            {
                string m_SQL = "SELECT * FROM [dbo].[EmployeeInfo]   ";
                using (SqlCommand m_cmd = g_DC.GetQueryCommand(p_EmployeeEntity, m_SQL, m_EmployeeInfo, m_objBetween))
                {
                    m_TransResult = g_DC.GetDataTable(p_EmployeeEntity, m_cmd);
                }
            }
            catch (Exception ex)
            {
                m_TransResult.LogMessage = ex.Message;
                m_TransResult.isSuccess = false;
            }
            return m_TransResult;
        }

        public SysEntity.TransResult GetEmployeeInfoGrid(SysEntity.Employee p_EmployeeEntity, Dictionary<string, Object> p_dicEmployeeInfo, string p_PageIndex, string p_PageSize, Dictionary<string, Object> p_Param)
        {
            SysEntity.TransResult m_TransResult = new SysEntity.TransResult();
            Dictionary<string, Object> m_objBetween = new Dictionary<string, Object>();

            string m_ChkSQL = "SELECT WorkID, isAdmin FROM  EmployeeInfo WHERE isDel='N' and isAdmin='Y' and  WorkID = '" + p_EmployeeEntity.WorkID+"'";
            using (SqlCommand m_cmd = new SqlCommand(m_ChkSQL))
            {

                DataTable m_DataTable = g_DC.GetDataTable(m_cmd);
                if (m_DataTable.Rows.Count != 1)
                {
                    p_dicEmployeeInfo.Add("CompanyCode", p_EmployeeEntity.CompanyCode);
                    p_dicEmployeeInfo.Add("Owenr", p_EmployeeEntity.WorkID);
                }
            }



            string[] m_LikeProp = { "DisplayName", "AreaName" };

            SysEntity.EmployeeInfo m_EmployeeInfo = g_FH.ConvertEntity<SysEntity.EmployeeInfo>(p_dicEmployeeInfo, ref m_objBetween);
            ArrayList m_alGrid = new ArrayList();

            try
            {
                string m_SQL = "SELECT [WorkID],[DisplayName],[AreaName],[CreateUser],[CompanyCode],[Owenr] FROM [dbo].[EmployeeInfo]   ";
                using (SqlCommand m_cmd = g_DC.GetQueryCommand(p_EmployeeEntity, m_SQL, m_EmployeeInfo, m_objBetween, m_LikeProp))
                {

                    m_TransResult = g_DC.GetDataTableByGrid(p_EmployeeEntity, m_cmd, p_Param, p_PageSize, p_PageIndex);
                    if (m_TransResult.isSuccess)
                    {
                        m_TransResult = g_DC.GridDataCtrl(m_TransResult, p_PageSize, p_Param);
                    }
                }
            }
            catch (Exception ex)
            {
                m_TransResult.LogMessage = ex.Message;
                m_TransResult.isSuccess = false;
            }
            return m_TransResult;
        }

        public SysEntity.TransResult AddEmployeeInfo(SysEntity.Employee p_EmployeeEntity, Dictionary<string, Object> p_dicEmployeeInfo)
        {
            SysEntity.TransResult m_TransResult = new SysEntity.TransResult();

            try
            {
                if (p_EmployeeEntity.CompanyCode != "1000")
                {
                    p_dicEmployeeInfo.Add("Password", "@123456");
                }

                m_TransResult = g_DC.InsertIntoByEntity<SysEntity.EmployeeInfo>(p_EmployeeEntity, "EmployeeInfo", p_dicEmployeeInfo);
                string m_SubGroupID = "";
                if (m_TransResult.isSuccess)
                {
                    Dictionary<string, Object> p_GetEmployeeInfo=new Dictionary<string, Object>();
                    p_GetEmployeeInfo.Add("WorkID", p_EmployeeEntity.WorkID);

                    m_TransResult = GetEmployeeInfos(p_EmployeeEntity, p_GetEmployeeInfo);
                    if (m_TransResult.isSuccess)
                    {
                        DataTable dt = (DataTable)m_TransResult.ResultEntity;
                        foreach (DataRow dr in dt.Rows)
                        {
                            m_SubGroupID = dr["SubGroupID"].ToString();
                        }
                        /*新增成功-用SubGroupID新增腳色*/
                        string m_SQL = "insert into EmployeeGroup (GroupID,WorkID,CreateUser,ModifyUser) values(@GroupID,@WorkID,@CreateUser,@ModifyUser)  ";
                        using (SqlCommand m_cmd = new SqlCommand(m_SQL))
                        {
                            m_cmd.Parameters.Add(new SqlParameter("GroupID", m_SubGroupID));
                            m_cmd.Parameters.Add(new SqlParameter("WorkID", p_dicEmployeeInfo["WorkID"]));
                            m_cmd.Parameters.Add(new SqlParameter("CreateUser", p_EmployeeEntity.WorkID));
                            m_cmd.Parameters.Add(new SqlParameter("ModifyUser", p_EmployeeEntity.WorkID));
                            g_DC.ExecuteNonQuery(p_EmployeeEntity, m_cmd);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                m_TransResult.LogMessage = ex.Message;
                m_TransResult.isSuccess = false;
            }

            return m_TransResult;

        }

        public SysEntity.TransResult DelEmployeeInfo(SysEntity.Employee p_EmployeeEntity, Dictionary<string, Object> p_dicEmployeeInfo)
        {
            SysEntity.TransResult m_TransResult = new SysEntity.TransResult();
            Dictionary<string, Object> m_ArrayList = new Dictionary<string, Object>();

            SysEntity.EmployeeInfo m_EmployeeInfo = g_FH.ConvertEntity<SysEntity.EmployeeInfo>(p_dicEmployeeInfo, ref m_ArrayList);

            try
            {
                string m_SQL = "Delete [dbo].[EmployeeInfo] Where [WorkID]=@WorkID";
                using (SqlCommand m_cmd = new SqlCommand(m_SQL))
                {
                    m_cmd.Parameters.Add(new SqlParameter("WorkID", m_EmployeeInfo.WorkID));
                    m_TransResult = g_DC.ExecuteNonQuery(p_EmployeeEntity, m_cmd);
                }
            }
            catch (Exception ex)
            {
                m_TransResult.LogMessage = ex.Message;
                m_TransResult.isSuccess = false;
            }

            return m_TransResult;

        }

        public SysEntity.TransResult UpdEmployeeInfo(SysEntity.Employee p_EmployeeEntity, Dictionary<string, Object> p_dicEmployeeInfo)
        {
            SysEntity.TransResult m_TransResult = new SysEntity.TransResult();
            Dictionary<string, Object> m_ArrayList = new Dictionary<string, Object>();

            SysEntity.EmployeeInfo m_EmployeeInfo = g_FH.ConvertEntity<SysEntity.EmployeeInfo>(p_dicEmployeeInfo, ref m_ArrayList);

            try
            {
                string m_SQL = "";
                string m_SQL_UPD = "";

                using (SqlCommand m_cmd = new SqlCommand())
                {
                    if (!string.IsNullOrEmpty( m_EmployeeInfo.GroupID))
                    {
                        m_SQL_UPD += " ,GroupID =@GroupID ";
                        m_cmd.Parameters.Add(new SqlParameter("GroupID", m_EmployeeInfo.GroupID));
                    }
                    if (!string.IsNullOrEmpty(m_EmployeeInfo.AreaName))
                    {
                        m_SQL_UPD += " ,AreaName =@AreaName ";
                        m_cmd.Parameters.Add(new SqlParameter("AreaName", m_EmployeeInfo.AreaName));
                    }
                    if (!string.IsNullOrEmpty(m_EmployeeInfo.Owenr))
                    {
                        m_SQL_UPD += " ,Owenr =@Owenr ";
                        m_cmd.Parameters.Add(new SqlParameter("Owenr", m_EmployeeInfo.Owenr));
                    }
                    if (!string.IsNullOrEmpty(m_EmployeeInfo.Password))
                    {
                        m_SQL_UPD += " ,Password =@Password ";
                        m_cmd.Parameters.Add(new SqlParameter("Password", m_EmployeeInfo.Password));
                    }
                    if (!string.IsNullOrEmpty(m_EmployeeInfo.CompanyCode))
                    {
                        m_SQL_UPD += " ,CompanyCode =@CompanyCode ";
                        m_cmd.Parameters.Add(new SqlParameter("CompanyCode", m_EmployeeInfo.CompanyCode));
                    }

                    

                    m_SQL = "Update [dbo].[EmployeeInfo]  set DisplayName =@DisplayName "+ m_SQL_UPD  + ",Remark =@Remark ,ModifyUser=@ModifyUser , ModifyDate=(Convert(varchar(10),Getdate(),111)), ModifyTime=(Convert(varchar(8),Getdate(),108)) ";
                    m_cmd.Parameters.Add(new SqlParameter("DisplayName", m_EmployeeInfo.DisplayName));
                    m_cmd.Parameters.Add(new SqlParameter("Remark", m_EmployeeInfo.Remark));
                    m_cmd.Parameters.Add(new SqlParameter("ModifyUser", m_EmployeeInfo.WorkID));


                   


                    if (m_EmployeeInfo.WorkID != "")
                    {
                        m_SQL += " where [WorkID]=@WorkID ";
                        m_cmd.Parameters.Add(new SqlParameter("WorkID", m_EmployeeInfo.WorkID));
                        m_cmd.CommandText = m_SQL;
                        m_TransResult = g_DC.ExecuteNonQuery(p_EmployeeEntity, m_cmd);
                    }
                    else
                    {
                        m_TransResult.LogMessage = "登入帳號 不可為空!!!";
                        m_TransResult.isSuccess = false;
                    }

                }
            }
            catch (Exception ex)
            {
                m_TransResult.LogMessage = ex.Message;
                m_TransResult.isSuccess = false;
            }

            return m_TransResult;

        }

        public SysEntity.TransResult GetEmployeeInfo(SysEntity.Employee p_EmployeeEntity, Dictionary<string, Object> p_dicEmployeeInfo)
        {
            SysEntity.TransResult m_TransResult = new SysEntity.TransResult();
            Dictionary<string, Object> m_objBetween = new Dictionary<string, Object>();
            SysEntity.ViewDetail m_ViewDetail = g_FH.ConvertEntity<SysEntity.ViewDetail>(p_dicEmployeeInfo, ref m_objBetween);

            try
            {
                string m_SQL = "SELECT   [WorkID],[DisplayName],[AreaName] from dbo.[EmployeeInfo]   where WorkID=@WorkID  ";
                p_dicEmployeeInfo = g_FH.GetKeyByDic("WorkID", p_dicEmployeeInfo);
                m_ViewDetail = g_FH.ConvertEntity<SysEntity.ViewDetail>(p_dicEmployeeInfo, ref m_objBetween);
                using (SqlCommand m_cmd = g_DC.GetQueryCommand(p_EmployeeEntity, m_SQL, m_ViewDetail, m_objBetween))
                {
                    m_TransResult = g_DC.GetDataTable(p_EmployeeEntity, m_cmd);
                }
            }
            catch (Exception ex)
            {
                m_TransResult.LogMessage = ex.Message;
                m_TransResult.isSuccess = false;
            }
            return m_TransResult;
        }

        public SysEntity.TransResult getEmployeeList_QRY(SysEntity.Employee p_EmployeeEntity, Dictionary<string, Object> p_dicParameter, string p_PageIndex, string p_PageSize, Dictionary<string, Object> p_Param, Boolean p_isDataOnly = false)
        {
            SysEntity.TransResult m_TransResult = new SysEntity.TransResult();
            Dictionary<string, Object> m_objBetween = new Dictionary<string, Object>();

            try
            {
                string m_SQL = " select from dbo.EmployeeInfo ";
                string[] m_LikeProp = { "DisplayName", "AreaName" };
                if (p_Param == null)
                {
                    p_Param = new Dictionary<string, Object>();
                    p_Param["OrderBy"] = "WorkID";
                }
                if (p_isDataOnly)
                {
                    m_TransResult = GetDataResult(p_EmployeeEntity, m_SQL, p_dicParameter);
                }
                else
                {

                    using (SqlCommand m_cmd = g_DC.GetQueryCommand(p_EmployeeEntity, m_SQL, p_dicParameter, m_objBetween, m_LikeProp))
                    {
                        m_TransResult = g_DC.GetDataTableByGrid(p_EmployeeEntity, m_cmd, p_Param, p_PageSize, p_PageIndex);
                        if (m_TransResult.isSuccess)
                        {
                            m_TransResult = g_DC.GridDataCtrl(m_TransResult, p_PageSize, p_Param);
                        }
                    }
                }

            }
            catch (Exception ex)
            {
                m_TransResult.LogMessage = ex.Message;
                m_TransResult.isSuccess = false;
            }
            return m_TransResult;
        }

        public SysEntity.TransResult GetDataResult(SysEntity.Employee p_EmployeeEntity, string p_SQL, Dictionary<string, Object> p_dicEntrty)
        {
            SysEntity.TransResult m_TransResult = new SysEntity.TransResult();
            string m_SQL = "select * from ( " + p_SQL + ") F  ";

            Int32 m_PararCount = 1;
            foreach (KeyValuePair<string, Object> item in p_dicEntrty)
            {
                if (m_PararCount == 1)
                {
                    m_SQL += " where " + item.Key + "=@" + item.Key;
                }
                else
                {
                    m_SQL += " and " + item.Key + "=@" + item.Key;

                }

                m_PararCount++;
            }
            using (SqlCommand m_cmd = new SqlCommand(m_SQL))
            {
                foreach (KeyValuePair<string, Object> item in p_dicEntrty)
                {
                    m_cmd.Parameters.Add(new SqlParameter(item.Key, item.Value));
                }

                m_TransResult = g_DC.GetDataTable(p_EmployeeEntity, m_cmd);
            }
            return m_TransResult;
        }

        #endregion

        #region FormInfo Methods

        public SysEntity.TransResult GetFormInfos(SysEntity.Employee p_EmployeeEntity, Dictionary<string, Object> p_dicFormInfo)
        {
            SysEntity.TransResult m_TransResult = new SysEntity.TransResult();
            Dictionary<string, Object> m_objBetween = new Dictionary<string, Object>();
            SysEntity.FormInfo m_FormInfo = g_FH.ConvertEntity<SysEntity.FormInfo>(p_dicFormInfo, ref m_objBetween);

            try
            {
                string m_SQL = "SELECT * FROM [dbo].[FormInfo]   ";
                using (SqlCommand m_cmd = g_DC.GetQueryCommand(p_EmployeeEntity, m_SQL, m_FormInfo, m_objBetween))
                {
                    m_TransResult = g_DC.GetDataTable(p_EmployeeEntity, m_cmd);
                }
            }
            catch (Exception ex)
            {
                m_TransResult.LogMessage = ex.Message;
                m_TransResult.isSuccess = false;
            }
            return m_TransResult;
        }

        public SysEntity.TransResult GetFormInfoGrid(SysEntity.Employee p_EmployeeEntity, Dictionary<string, Object> p_dicFormInfo, string p_PageIndex, string p_PageSize, Dictionary<string, Object> p_Param)
        {
            SysEntity.TransResult m_TransResult = new SysEntity.TransResult();
            Dictionary<string, Object> m_objBetween = new Dictionary<string, Object>();
            SysEntity.FormInfo m_FormInfo = g_FH.ConvertEntity<SysEntity.FormInfo>(p_dicFormInfo, ref m_objBetween);
            ArrayList m_alGrid = new ArrayList();

            string[] m_LikeProp = { "Description" };
            try
            {
                string m_SQL = "SELECT [FormID],[Description],[CreateUser] FROM [dbo].[FormInfo]   ";
                using (SqlCommand m_cmd = g_DC.GetQueryCommand(p_EmployeeEntity, m_SQL, m_FormInfo, m_objBetween, m_LikeProp))
                {

                    m_TransResult = g_DC.GetDataTableByGrid(p_EmployeeEntity, m_cmd, p_Param, p_PageSize, p_PageIndex);
                    if (m_TransResult.isSuccess)
                    {
                        m_TransResult = g_DC.GridDataCtrl(m_TransResult, p_PageSize, p_Param);
                    }
                }
            }
            catch (Exception ex)
            {
                m_TransResult.LogMessage = ex.Message;
                m_TransResult.isSuccess = false;
            }
            return m_TransResult;
        }

        public SysEntity.TransResult AddFormInfo(SysEntity.Employee p_EmployeeEntity, Dictionary<string, Object> p_dicFormInfo)
        {
            SysEntity.TransResult m_TransResult = new SysEntity.TransResult();

            try
            {
                m_TransResult = g_DC.InsertIntoByEntity<SysEntity.FormInfo>(p_EmployeeEntity, "FormInfo", p_dicFormInfo);
            }
            catch (Exception ex)
            {
                m_TransResult.LogMessage = ex.Message;
                m_TransResult.isSuccess = false;
            }

            return m_TransResult;

        }

        public SysEntity.TransResult DelFormInfo(SysEntity.Employee p_EmployeeEntity, Dictionary<string, Object> p_dicFormInfo)
        {
            SysEntity.TransResult m_TransResult = new SysEntity.TransResult();
            Dictionary<string, Object> m_ArrayList = new Dictionary<string, Object>();

            SysEntity.FormInfo m_FormInfo = g_FH.ConvertEntity<SysEntity.FormInfo>(p_dicFormInfo, ref m_ArrayList);

            try
            {
                string m_SQL = "Delete [dbo].[FormInfo] Where [FormID]=@FormID";
                using (SqlCommand m_cmd = new SqlCommand(m_SQL))
                {
                    m_cmd.Parameters.Add(new SqlParameter("FormID", m_FormInfo.FormID));
                    m_TransResult = g_DC.ExecuteNonQuery(p_EmployeeEntity, m_cmd);
                }
            }
            catch (Exception ex)
            {
                m_TransResult.LogMessage = ex.Message;
                m_TransResult.isSuccess = false;
            }

            return m_TransResult;

        }

        public SysEntity.TransResult UpdFormInfo(SysEntity.Employee p_EmployeeEntity, Dictionary<string, Object> p_dicFormInfo)
        {
            SysEntity.TransResult m_TransResult = new SysEntity.TransResult();
            Dictionary<string, Object> m_ArrayList = new Dictionary<string, Object>();

            SysEntity.FormInfo m_FormInfo = g_FH.ConvertEntity<SysEntity.FormInfo>(p_dicFormInfo, ref m_ArrayList);

            try
            {
                string m_SQL = "";
                using (SqlCommand m_cmd = new SqlCommand())
                {
                    m_SQL = "Update [dbo].[FormInfo]  set Path =@Path, Description =@Description ,ModifyUser=@ModifyUser , ModifyDate=(Convert(varchar(10),Getdate(),111)), ModifyTime=(Convert(varchar(8),Getdate(),108)) ";
                 
                    m_cmd.Parameters.Add(new SqlParameter("Path", m_FormInfo.Path));
                    m_cmd.Parameters.Add(new SqlParameter("Description", m_FormInfo.Description));
                    m_cmd.Parameters.Add(new SqlParameter("ModifyUser", p_EmployeeEntity.WorkID));

                    if (m_FormInfo.FormID != "")
                    {
                        m_SQL += " where [FormID]=@FormID ";
                        m_cmd.Parameters.Add(new SqlParameter("FormID", m_FormInfo.FormID));
                        m_cmd.CommandText = m_SQL;
                        m_TransResult = g_DC.ExecuteNonQuery(p_EmployeeEntity, m_cmd);
                    }
                    else
                    {
                        m_TransResult.LogMessage = "FormID 不可為空!!!";
                        m_TransResult.isSuccess = false;
                    }

                }
            }
            catch (Exception ex)
            {
                m_TransResult.LogMessage = ex.Message;
                m_TransResult.isSuccess = false;
            }

            return m_TransResult;

        }

        #endregion

        #region ModuleInfo Methods


        public SysEntity.TransResult GetModuleInfos(SysEntity.Employee p_EmployeeEntity, Dictionary<string, Object> p_dicModuleInfo)
        {
            SysEntity.TransResult m_TransResult = new SysEntity.TransResult();
            Dictionary<string, Object> m_objBetween = new Dictionary<string, Object>();
            SysEntity.ModuleInfo m_ModuleInfo = g_FH.ConvertEntity<SysEntity.ModuleInfo>(p_dicModuleInfo, ref m_objBetween);

            try
            {
                string m_SQL = "SELECT * FROM [dbo].[ModuleInfo]   ";
                using (SqlCommand m_cmd = g_DC.GetQueryCommand(p_EmployeeEntity, m_SQL, m_ModuleInfo, m_objBetween))
                {
                    m_TransResult = g_DC.GetDataTable(p_EmployeeEntity, m_cmd);
                }
            }
            catch (Exception ex)
            {
                m_TransResult.LogMessage = ex.Message;
                m_TransResult.isSuccess = false;
            }
            return m_TransResult;
        }

        public SysEntity.TransResult GetModuleInfoGrid(SysEntity.Employee p_EmployeeEntity, Dictionary<string, Object> p_dicModuleInfo, string p_PageIndex, string p_PageSize, Dictionary<string, Object> p_Param)
        {
            SysEntity.TransResult m_TransResult = new SysEntity.TransResult();
            Dictionary<string, Object> m_objBetween = new Dictionary<string, Object>();
            SysEntity.ModuleInfo m_ModuleInfo = g_FH.ConvertEntity<SysEntity.ModuleInfo>(p_dicModuleInfo, ref m_objBetween);
            ArrayList m_alGrid = new ArrayList();


            try
            {
                string m_SQL = "SELECT [ModuleKey],[ModuleName],[CreateUser] FROM [dbo].[ModuleInfo]   ";
                using (SqlCommand m_cmd = g_DC.GetQueryCommand(p_EmployeeEntity, m_SQL, m_ModuleInfo, m_objBetween))
                {

                    m_TransResult = g_DC.GetDataTableByGrid(p_EmployeeEntity, m_cmd, p_Param, p_PageSize, p_PageIndex);
                    if (m_TransResult.isSuccess)
                    {
                        m_TransResult = g_DC.GridDataCtrl(m_TransResult, p_PageSize, p_Param);
                    }
                }
            }
            catch (Exception ex)
            {
                m_TransResult.LogMessage = ex.Message;
                m_TransResult.isSuccess = false;
            }
            return m_TransResult;
        }

        public SysEntity.TransResult AddModuleInfo(SysEntity.Employee p_EmployeeEntity, Dictionary<string, Object> p_dicModuleInfo)
        {
            SysEntity.TransResult m_TransResult = new SysEntity.TransResult();

            try
            {
                m_TransResult = g_DC.InsertIntoByEntity<SysEntity.ModuleInfo>(p_EmployeeEntity, "ModuleInfo", p_dicModuleInfo);
            }
            catch (Exception ex)
            {
                m_TransResult.LogMessage = ex.Message;
                m_TransResult.isSuccess = false;
            }

            return m_TransResult;

        }

        public SysEntity.TransResult DelModuleInfo(SysEntity.Employee p_EmployeeEntity, Dictionary<string, Object> p_dicModuleInfo)
        {
            SysEntity.TransResult m_TransResult = new SysEntity.TransResult();
            Dictionary<string, Object> m_ArrayList = new Dictionary<string, Object>();

            SysEntity.ModuleInfo m_ModuleInfo = g_FH.ConvertEntity<SysEntity.ModuleInfo>(p_dicModuleInfo, ref m_ArrayList);

            try
            {
                string m_SQL = "Delete [dbo].[ModuleInfo] Where [ModuleKey]=@ModuleKey";
                using (SqlCommand m_cmd = new SqlCommand(m_SQL))
                {
                    m_cmd.Parameters.Add(new SqlParameter("ModuleKey", m_ModuleInfo.ModuleKey));
                    m_TransResult = g_DC.ExecuteNonQuery(p_EmployeeEntity, m_cmd);
                }
            }
            catch (Exception ex)
            {
                m_TransResult.LogMessage = ex.Message;
                m_TransResult.isSuccess = false;
            }

            return m_TransResult;

        }

        public SysEntity.TransResult UpdModuleInfo(SysEntity.Employee p_EmployeeEntity, Dictionary<string, Object> p_dicModuleInfo)
        {
            SysEntity.TransResult m_TransResult = new SysEntity.TransResult();
            Dictionary<string, Object> m_ArrayList = new Dictionary<string, Object>();

            SysEntity.ModuleInfo m_ModuleInfo = g_FH.ConvertEntity<SysEntity.ModuleInfo>(p_dicModuleInfo, ref m_ArrayList);

            try
            {
                string m_SQL = "";
                using (SqlCommand m_cmd = new SqlCommand())
                {
                    m_SQL = "Update [dbo].[ModuleInfo]  set ModuleName =@ModuleName, IconCSS =@IconCSS, ModulePagedep =@ModulePagedep,ModuleContent=@ModuleContent , ModuleImg =@ModuleImg ,ModifyUser=@ModifyUser , ModifyDate=(Convert(varchar(10),Getdate(),111)), ModifyTime=(Convert(varchar(8),Getdate(),108)) ";
                    m_cmd.Parameters.Add(new SqlParameter("ModuleName", m_ModuleInfo.ModuleName));
                    m_cmd.Parameters.Add(new SqlParameter("IconCSS", m_ModuleInfo.IconCSS));
                    m_cmd.Parameters.Add(new SqlParameter("ModulePagedep", m_ModuleInfo.ModulePagedep));
                    m_cmd.Parameters.Add(new SqlParameter("ModuleContent", m_ModuleInfo.ModuleContent));
                    m_cmd.Parameters.Add(new SqlParameter("ModuleImg", m_ModuleInfo.ModuleImg));
                    m_cmd.Parameters.Add(new SqlParameter("ModifyUser", p_EmployeeEntity.WorkID));

                    if (m_ModuleInfo.ModuleKey != "")
                    {
                        m_SQL += " where [ModuleKey]=@ModuleKey ";
                        m_cmd.Parameters.Add(new SqlParameter("ModuleKey", m_ModuleInfo.ModuleKey));
                        m_cmd.CommandText = m_SQL;
                        m_TransResult = g_DC.ExecuteNonQuery(p_EmployeeEntity, m_cmd);
                    }
                    else
                    {
                        m_TransResult.LogMessage = "ModuleKey 不可為空!!!";
                        m_TransResult.isSuccess = false;
                    }

                }
            }
            catch (Exception ex)
            {
                m_TransResult.LogMessage = ex.Message;
                m_TransResult.isSuccess = false;
            }

            return m_TransResult;

        }

        public DataTable GetModuleInfoByMenuID(SysEntity.Employee p_EmployeeEntity, string p_MenuID)
        {
            DataTable m_dt = new DataTable();

            string m_SQL = "  select D.*, Description MenuDesc From [dbo].[ModuleInfo] D  Left Join MainMenu M  on D.ModuleKey=M.ModuleKey   Left Join FormInfo F  on M.FormID=F.FormID Where MenuID=@MenuID  ";


            using (SqlCommand m_cmd = new SqlCommand(m_SQL))
            {
                m_cmd.Parameters.Add(new SqlParameter("MenuID", p_MenuID));


                SysEntity.TransResult m_TransResult = new SysEntity.TransResult();

                m_TransResult = g_DC.GetDataTable(p_EmployeeEntity, m_cmd);
                if (m_TransResult.isSuccess)
                {
                    m_dt = (DataTable)m_TransResult.ResultEntity;
                }
            }
            return m_dt;
        }




        #endregion

        #region Form Methods



        public SysEntity.TransResult GetMenuForm(SysEntity.Employee p_EmployeeEntity)
        {
            SysEntity.TransResult m_TransResult = new SysEntity.TransResult();
            try
            {
                using (SqlCommand m_cmd = new SqlCommand(" select [FormID],[Description] from [FormInfo]  "))
                {
                    var m_Params = new
                    {
                      
                    };

                    m_TransResult = g_DC.GetData<SysEntity.FormInfo>(p_EmployeeEntity, m_cmd, m_Params);
                    if (m_TransResult.isSuccess)
                    {
                        List<SysEntity.FormInfo> m_lisFormInfo = (List<SysEntity.FormInfo>)m_TransResult.ResultEntity;

                        if (m_lisFormInfo.Count != 0)
                        {
                            m_TransResult.ResultEntity = m_lisFormInfo;
                        }
                        else
                        {
                            m_TransResult.isSuccess = false;
                            m_TransResult.LogMessage = "FormID,data is empty";
                        }

                    }
                }
            }
            catch (Exception ex)
            {
                m_TransResult.LogMessage = ex.Message;
                m_TransResult.isSuccess = false;
            }
            return m_TransResult;
        }





        public SysEntity.TransResult GetForms(SysEntity.Employee p_EmployeeEntity, Dictionary<string, Object> p_dicForm)
        {
            SysEntity.TransResult m_TransResult = new SysEntity.TransResult();
            Dictionary<string, Object> m_objBetween = new Dictionary<string, Object>();
            SysEntity.Form m_Form = g_FH.ConvertEntity<SysEntity.Form>(p_dicForm, ref m_objBetween);

            try
            {
                string m_SQL = "SELECT * FROM [dbo].[Form]   ";
                using (SqlCommand m_cmd = g_DC.GetQueryCommand(p_EmployeeEntity, m_SQL, m_Form, m_objBetween))
                {
                    m_TransResult = g_DC.GetDataTable(p_EmployeeEntity, m_cmd);
                }
            }
            catch (Exception ex)
            {
                m_TransResult.LogMessage = ex.Message;
                m_TransResult.isSuccess = false;
            }
            return m_TransResult;
        }

        public SysEntity.TransResult GetFormGrid(SysEntity.Employee p_EmployeeEntity, Dictionary<string, Object> p_dicForm, string p_PageIndex, string p_PageSize, Dictionary<string, Object> p_Param)
        {
            SysEntity.TransResult m_TransResult = new SysEntity.TransResult();
            Dictionary<string, Object> m_objBetween = new Dictionary<string, Object>();
            SysEntity.Form m_Form = g_FH.ConvertEntity<SysEntity.Form>(p_dicForm, ref m_objBetween);
            ArrayList m_alGrid = new ArrayList();


            try
            {
                string m_SQL = "SELECT [FormID],[Description],[CreateUser],[ModifyUser],[CreateDate],[ModifyDate] FROM [dbo].[Form]   ";
                using (SqlCommand m_cmd = g_DC.GetQueryCommand(p_EmployeeEntity, m_SQL, m_Form, m_objBetween))
                {

                    m_TransResult = g_DC.GetDataTableByGrid(p_EmployeeEntity, m_cmd, p_Param, p_PageSize, p_PageIndex);
                    if (m_TransResult.isSuccess)
                    {
                        m_TransResult = g_DC.GridDataCtrl(m_TransResult, p_PageSize, p_Param);
                    }
                }
            }
            catch (Exception ex)
            {
                m_TransResult.LogMessage = ex.Message;
                m_TransResult.isSuccess = false;
            }
            return m_TransResult;
        }

        public SysEntity.TransResult AddForm(SysEntity.Employee p_EmployeeEntity, Dictionary<string, Object> p_dicForm)
        {
            SysEntity.TransResult m_TransResult = new SysEntity.TransResult();

            try
            {
                m_TransResult = g_DC.InsertIntoByEntity<SysEntity.Form>(p_EmployeeEntity, "Form", p_dicForm);
            }
            catch (Exception ex)
            {
                m_TransResult.LogMessage = ex.Message;
                m_TransResult.isSuccess = false;
            }

            return m_TransResult;

        }

        public SysEntity.TransResult DelForm(SysEntity.Employee p_EmployeeEntity, Dictionary<string, Object> p_dicForm)
        {
            SysEntity.TransResult m_TransResult = new SysEntity.TransResult();
            Dictionary<string, Object> m_ArrayList = new Dictionary<string, Object>();

            SysEntity.Form m_Form = g_FH.ConvertEntity<SysEntity.Form>(p_dicForm, ref m_ArrayList);

            try
            {
                string m_SQL = "Delete [dbo].[Form] Where [FormID]=@FormID";
                using (SqlCommand m_cmd = new SqlCommand(m_SQL))
                {
                    m_cmd.Parameters.Add(new SqlParameter("FormID", m_Form.FormID));
                    m_TransResult = g_DC.ExecuteNonQuery(p_EmployeeEntity, m_cmd);
                }
            }
            catch (Exception ex)
            {
                m_TransResult.LogMessage = ex.Message;
                m_TransResult.isSuccess = false;
            }

            return m_TransResult;

        }

        public SysEntity.TransResult UpdForm(SysEntity.Employee p_EmployeeEntity, Dictionary<string, Object> p_dicForm)
        {
            SysEntity.TransResult m_TransResult = new SysEntity.TransResult();
            Dictionary<string, Object> m_ArrayList = new Dictionary<string, Object>();

            SysEntity.Form m_Form = g_FH.ConvertEntity<SysEntity.Form>(p_dicForm, ref m_ArrayList);

            try
            {
                string m_SQL = "";
                using (SqlCommand m_cmd = new SqlCommand())
                {
                    m_SQL = "Update [dbo].[Form]  set Description =@Description , ModifyUser=@ModifyUser , ModifyDate=(Convert(varchar(10),Getdate(),111)), ModifyTime=(Convert(varchar(8),Getdate(),108)) ";
                    m_cmd.Parameters.Add(new SqlParameter("Description", m_Form.Description));
          
                    m_cmd.Parameters.Add(new SqlParameter("ModifyUser", p_EmployeeEntity.WorkID));

                    if (m_Form.FormID != "")
                    {
                        m_SQL += " where [FormID]=@FormID ";
                        m_cmd.Parameters.Add(new SqlParameter("FormID", m_Form.FormID));
                        m_cmd.CommandText = m_SQL;
                        m_TransResult = g_DC.ExecuteNonQuery(p_EmployeeEntity, m_cmd);
                    }
                    else
                    {
                        m_TransResult.LogMessage = "FormID 不可為空!!!";
                        m_TransResult.isSuccess = false;
                    }

                }
            }
            catch (Exception ex)
            {
                m_TransResult.LogMessage = ex.Message;
                m_TransResult.isSuccess = false;
            }

            return m_TransResult;

        }

        #endregion

        #region Site Methods

        public SysEntity.TransResult GetSites(SysEntity.Employee p_EmployeeEntity, Dictionary<string, Object> p_dicSite)
        {
            SysEntity.TransResult m_TransResult = new SysEntity.TransResult();
            Dictionary<string, Object> m_objBetween = new Dictionary<string, Object>();
            SysEntity.Site m_Site = g_FH.ConvertEntity<SysEntity.Site>(p_dicSite, ref m_objBetween);

            try
            {
                string m_SQL = "SELECT [SiteID],[Language],[Currency],[Description],[ModifyUser],[ModifyDate] FROM [dbo].[Site]   ";
                using (SqlCommand m_cmd = g_DC.GetQueryCommand(p_EmployeeEntity, m_SQL, m_Site, m_objBetween))
                {
                    m_TransResult = g_DC.GetDataTable(p_EmployeeEntity, m_cmd);
                }
            }
            catch (Exception ex)
            {
                m_TransResult.LogMessage = ex.Message;
                m_TransResult.isSuccess = false;
            }
            return m_TransResult;
        }

        public SysEntity.TransResult GetSiteGrid(SysEntity.Employee p_EmployeeEntity, Dictionary<string, Object> p_dicSite, string p_PageIndex, string p_PageSize, Dictionary<string, Object> p_Param)
        {

            SysEntity.TransResult m_TransResult = new SysEntity.TransResult();
            Dictionary<string, Object> m_objBetween = new Dictionary<string, Object>();
            SysEntity.Site m_Site = g_FH.ConvertEntity<SysEntity.Site>(p_dicSite, ref m_objBetween);
            ArrayList m_alGrid = new ArrayList();


            try
            {
                string m_SQL = "SELECT [SiteID],[Language],[Currency],[Description],[ModifyUser],[ModifyDate] FROM [dbo].[Site]    ";
                using (SqlCommand m_cmd = g_DC.GetQueryCommand(p_EmployeeEntity, m_SQL, m_Site, m_objBetween))
                {

                    m_TransResult = g_DC.GetDataTableByGrid(p_EmployeeEntity, m_cmd, p_Param, p_PageSize, p_PageIndex);
                    if (m_TransResult.isSuccess)
                    {
                        m_TransResult = g_DC.GridDataCtrl(m_TransResult, p_PageSize, p_Param);
                    }
                }
            }
            catch (Exception ex)
            {
                m_TransResult.LogMessage = ex.Message;
                m_TransResult.isSuccess = false;
            }
            return m_TransResult;
        }

        public SysEntity.TransResult AddSite(SysEntity.Employee p_EmployeeEntity, Dictionary<string, Object> p_dicSite)
        {
            SysEntity.TransResult m_TransResult = new SysEntity.TransResult();
            try
            {
                m_TransResult = g_DC.InsertIntoByEntity<SysEntity.Site>(p_EmployeeEntity, "Site", p_dicSite);
            }
            catch (Exception ex)
            {
                m_TransResult.LogMessage = ex.Message;
                m_TransResult.isSuccess = false;
            }

            return m_TransResult;

        }

        public SysEntity.TransResult DelSite(SysEntity.Employee p_EmployeeEntity, Dictionary<string, Object> p_dicSite)
        {
            SysEntity.TransResult m_TransResult = new SysEntity.TransResult();
            Dictionary<string, Object> m_ArrayList = new Dictionary<string, Object>();

            SysEntity.Site m_Site = g_FH.ConvertEntity<SysEntity.Site>(p_dicSite, ref m_ArrayList);

            try
            {
                string m_SQL = "Delete [dbo].[Site] Where [SiteID]=@SiteID";
                using (SqlCommand m_cmd = new SqlCommand(m_SQL))
                {
                    m_cmd.Parameters.Add(new SqlParameter("SiteID", m_Site.SiteID));
                    m_TransResult = g_DC.ExecuteNonQuery(p_EmployeeEntity, m_cmd);
                }
            }
            catch (Exception ex)
            {
                m_TransResult.LogMessage = ex.Message;
                m_TransResult.isSuccess = false;
            }

            return m_TransResult;

        }

        public SysEntity.TransResult UpdSite(SysEntity.Employee p_EmployeeEntity, Dictionary<string, Object> p_dicSite)
        {
            SysEntity.TransResult m_TransResult = new SysEntity.TransResult();
            Dictionary<string, Object> m_ArrayList = new Dictionary<string, Object>();

            SysEntity.Site m_Site = g_FH.ConvertEntity<SysEntity.Site>(p_dicSite, ref m_ArrayList);

            try
            {
                string m_SQL = "";
                using (SqlCommand m_cmd = new SqlCommand())
                {
                    m_SQL = "Update [dbo].[Site]  set Description =@Description , Language =@Language,Currency =@Currency,ModifyUser=@ModifyUser , ModifyDate=(Convert(varchar(10),Getdate(),111)), ModifyTime=(Convert(varchar(8),Getdate(),108)) ";
                    m_cmd.Parameters.Add(new SqlParameter("Description", m_Site.Description));
                    m_cmd.Parameters.Add(new SqlParameter("Currency", m_Site.Currency));
                    m_cmd.Parameters.Add(new SqlParameter("Language", m_Site.Language));
                    m_cmd.Parameters.Add(new SqlParameter("ModifyUser", p_EmployeeEntity.WorkID));

                    if (m_Site.SiteID != "")
                    {
                        m_SQL += " where [SiteID]=@SiteID ";
                        m_cmd.Parameters.Add(new SqlParameter("SiteID", m_Site.SiteID));
                        m_cmd.CommandText = m_SQL;
                        m_TransResult = g_DC.ExecuteNonQuery(p_EmployeeEntity, m_cmd);
                    }
                    else
                    {
                        m_TransResult.LogMessage = "SiteID 不可為空!!!";
                        m_TransResult.isSuccess = false;
                    }

                }
            }
            catch (Exception ex)
            {
                m_TransResult.LogMessage = ex.Message;
                m_TransResult.isSuccess = false;
            }

            return m_TransResult;

        }

        public SysEntity.TransResult GetSiteMenu(SysEntity.Employee p_EmployeeEntity)
        {
            SysEntity.TransResult m_TransResult = new SysEntity.TransResult();

            try
            {
                string m_SQL = "select M.SiteID,S.Description from 	(select distinct SiteID from [dbo].[MainMenu]) M left join (SELECT SiteID,DescriptionEnglish Description FROM [dbo].[Site] ) S on M.SiteID=S.SiteID order by SiteID   ";

                using (SqlCommand m_cmd = new SqlCommand(m_SQL))
                {
                    m_TransResult = g_DC.GetDataTable(p_EmployeeEntity, m_cmd);
                }

            }
            catch (Exception ex)
            {
                m_TransResult.LogMessage = ex.Message;
                m_TransResult.isSuccess = false;
            }
            return m_TransResult;
        }






        #endregion

        #region SiteForm Methods

        public SysEntity.TransResult GetSiteForms(SysEntity.Employee p_EmployeeEntity, Dictionary<string, Object> p_dicSiteForm)
        {
            SysEntity.TransResult m_TransResult = new SysEntity.TransResult();
            Dictionary<string, Object> m_objBetween = new Dictionary<string, Object>();
            SysEntity.SiteForm m_SiteForm = g_FH.ConvertEntity<SysEntity.SiteForm>(p_dicSiteForm, ref m_objBetween);

            try
            {
                string m_SQL = "SELECT [SiteFormID],[FormID],[SiteID],[Admin],[NotesPath],[AgentName],[ValidID],[DocLockTime],[Levels] FROM [dbo].[SiteForm]   ";
                using (SqlCommand m_cmd = g_DC.GetQueryCommand(p_EmployeeEntity, m_SQL, m_SiteForm, m_objBetween))
                {
                    m_TransResult = g_DC.GetDataTable(p_EmployeeEntity, m_cmd);
                }
            }
            catch (Exception ex)
            {
                m_TransResult.LogMessage = ex.Message;
                m_TransResult.isSuccess = false;
            }
            return m_TransResult;
        }

        public SysEntity.TransResult GetSiteFormGrid(SysEntity.Employee p_EmployeeEntity, Dictionary<string, Object> p_dicSiteForm, string p_PageIndex, string p_PageSize, Dictionary<string, Object> p_Param)
        {

            SysEntity.TransResult m_TransResult = new SysEntity.TransResult();
            Dictionary<string, Object> m_objBetween = new Dictionary<string, Object>();
            SysEntity.SiteForm m_SiteForm = g_FH.ConvertEntity<SysEntity.SiteForm>(p_dicSiteForm, ref m_objBetween);
            ArrayList m_alGrid = new ArrayList();


            try
            {
                string m_SQL = "SELECT [SiteFormID],[AgentName],[FormID],[SiteID],[Admin],[NotesPath] FROM [dbo].[SiteForm]    ";
                using (SqlCommand m_cmd = g_DC.GetQueryCommand(p_EmployeeEntity, m_SQL, m_SiteForm, m_objBetween))
                {

                    m_TransResult = g_DC.GetDataTableByGrid(p_EmployeeEntity, m_cmd, p_Param, p_PageSize, p_PageIndex);
                    if (m_TransResult.isSuccess)
                    {
                        m_TransResult = g_DC.GridDataCtrl(m_TransResult, p_PageSize, p_Param);
                    }
                }
            }
            catch (Exception ex)
            {
                m_TransResult.LogMessage = ex.Message;
                m_TransResult.isSuccess = false;
            }
            return m_TransResult;
        }

        public SysEntity.TransResult AddSiteForm(SysEntity.Employee p_EmployeeEntity, Dictionary<string, Object> p_dicSiteForm)
        {
            SysEntity.TransResult m_TransResult = new SysEntity.TransResult();

            try
            {
                m_TransResult = g_DC.InsertIntoByEntity<SysEntity.SiteForm>(p_EmployeeEntity, "SiteForm", p_dicSiteForm);
            }
            catch (Exception ex)
            {
                m_TransResult.LogMessage = ex.Message;
                m_TransResult.isSuccess = false;
            }

            return m_TransResult;

        }

        public SysEntity.TransResult DelSiteForm(SysEntity.Employee p_EmployeeEntity, Dictionary<string, Object> p_dicSiteForm)
        {
            SysEntity.TransResult m_TransResult = new SysEntity.TransResult();
            Dictionary<string, Object> m_ArrayList = new Dictionary<string, Object>();

            SysEntity.SiteForm m_SiteForm = g_FH.ConvertEntity<SysEntity.SiteForm>(p_dicSiteForm, ref m_ArrayList);

            try
            {
                string m_SQL = "Delete [dbo].[SiteForm] Where [SiteFormID]=@SiteFormID";
                using (SqlCommand m_cmd = new SqlCommand(m_SQL))
                {
                    m_cmd.Parameters.Add(new SqlParameter("SiteFormID", m_SiteForm.SiteID + m_SiteForm.FormID));
                    m_TransResult = g_DC.ExecuteNonQuery(p_EmployeeEntity, m_cmd);
                }
            }
            catch (Exception ex)
            {
                m_TransResult.LogMessage = ex.Message;
                m_TransResult.isSuccess = false;
            }

            return m_TransResult;

        }

        public SysEntity.TransResult UpdSiteForm(SysEntity.Employee p_EmployeeEntity, Dictionary<string, Object> p_dicSiteForm)
        {
            SysEntity.TransResult m_TransResult = new SysEntity.TransResult();
            Dictionary<string, Object> m_ArrayList = new Dictionary<string, Object>();

            SysEntity.SiteForm m_SiteForm = g_FH.ConvertEntity<SysEntity.SiteForm>(p_dicSiteForm, ref m_ArrayList);

            try
            {
                string m_SQL = "";
                using (SqlCommand m_cmd = new SqlCommand())
                {
                    m_SQL = "Update [dbo].[SiteForm] set Admin =@Admin, NotesPath=@NotesPath,AgentName=@AgentName,DocLockTime=@DocLockTime,{0} Levels=@Levels,ModifyUser=@ModifyUser , ModifyDate=(Convert(varchar(10),Getdate(),111)), ModifyTime=(Convert(varchar(8),Getdate(),108)) ";
                    m_cmd.Parameters.Add(new SqlParameter("Admin", m_SiteForm.Admin));
                    m_cmd.Parameters.Add(new SqlParameter("NotesPath", m_SiteForm.NotesPath));
                    m_cmd.Parameters.Add(new SqlParameter("AgentName", m_SiteForm.AgentName));
                    m_cmd.Parameters.Add(new SqlParameter("DocLockTime", m_SiteForm.DocLockTime));
                    if (m_SiteForm.ValidID == "")
                    {
                        m_SQL = string.Format(m_SQL, " ");
                    }
                    else
                    {
                        m_SQL = string.Format(m_SQL, " ValidID=@ValidID,");
                        m_cmd.Parameters.Add(new SqlParameter("ValidID", m_SiteForm.ValidID));

                    }
                    m_cmd.Parameters.Add(new SqlParameter("Levels", m_SiteForm.Levels));
                    m_cmd.Parameters.Add(new SqlParameter("ModifyUser", p_EmployeeEntity.WorkID));

                    if (m_SiteForm.SiteFormID != "")
                    {
                        m_SQL += " where [SiteFormID]=@SiteFormID and [FormID]=@FormID and [SiteID]=@SiteID  ";
                        m_cmd.Parameters.Add(new SqlParameter("SiteFormID", m_SiteForm.SiteID + m_SiteForm.FormID));
                        m_cmd.Parameters.Add(new SqlParameter("SiteID", m_SiteForm.SiteID));
                        m_cmd.Parameters.Add(new SqlParameter("FormID", m_SiteForm.FormID));

                        m_cmd.CommandText = m_SQL;
                        m_TransResult = g_DC.ExecuteNonQuery(p_EmployeeEntity, m_cmd);
                    }
                    else
                    {
                        m_TransResult.LogMessage = "SiteFormID 不可為空!!!";
                        m_TransResult.isSuccess = false;
                    }

                }
            }
            catch (Exception ex)
            {
                m_TransResult.LogMessage = ex.Message;
                m_TransResult.isSuccess = false;
            }

            return m_TransResult;

        }

        public SysEntity.TransResult GetSiteFormsWithoutSetting(SysEntity.Employee p_EmployeeEntity, Dictionary<string, Object> p_dicSiteForm)
        {
            SysEntity.TransResult m_TransResult = new SysEntity.TransResult();
            Dictionary<string, Object> m_objBetween = new Dictionary<string, Object>();
            SysEntity.SiteForm m_SiteForm = g_FH.ConvertEntity<SysEntity.SiteForm>(p_dicSiteForm, ref m_objBetween);

            try
            {
                string m_SQL = "SELECT [SiteFormID],[FormID],[SiteID],[Admin],[NotesPath],[AgentName],[ValidID],[DocLockTime],[Levels] FROM [dbo].[SiteForm] where [SiteFormID] not in(select distinct [SiteFormID] from  [dbo].[ViewDetail])   ";
                using (SqlCommand m_cmd = g_DC.GetQueryCommand(p_EmployeeEntity, m_SQL, m_SiteForm, m_objBetween))
                {
                    m_TransResult = g_DC.GetDataTable(p_EmployeeEntity, m_cmd);
                }
            }
            catch (Exception ex)
            {
                m_TransResult.LogMessage = ex.Message;
                m_TransResult.isSuccess = false;
            }
            return m_TransResult;
        }

        #endregion

        #region MultiLanguage

        public SysEntity.TransResult GetMultiLanguages(SysEntity.Employee p_EmployeeEntity, Dictionary<string, Object> p_dicMultiLanguage)
        {
            SysEntity.TransResult m_TransResult = new SysEntity.TransResult();
            Dictionary<string, Object> m_objBetween = new Dictionary<string, Object>();
            SysEntity.MultiLanguage m_MultiLanguage = g_FH.ConvertEntity<SysEntity.MultiLanguage>(p_dicMultiLanguage, ref m_objBetween);

            try
            {
                string m_SQL = "SELECT [LanguageKey],[ValueTW],[ValueCN],[ValueUS],[ValueJP],[ValueBR],[ModifyUser],[ModifyDate] FROM [dbo].[MultiLanguage]   ";
                using (SqlCommand m_cmd = g_DC.GetQueryCommand(p_EmployeeEntity, m_SQL, m_MultiLanguage, m_objBetween))
                {
                    m_TransResult = g_DC.GetDataTable(p_EmployeeEntity, m_cmd);
                }
            }
            catch (Exception ex)
            {
                m_TransResult.LogMessage = ex.Message;
                m_TransResult.isSuccess = false;
            }
            return m_TransResult;
        }

        public SysEntity.TransResult GetMultiLanguageGrid(SysEntity.Employee p_EmployeeEntity, Dictionary<string, Object> p_dicMultiLanguage, string p_PageIndex, string p_PageSize, Dictionary<string, Object> p_Param)
        {

            SysEntity.TransResult m_TransResult = new SysEntity.TransResult();
            Dictionary<string, Object> m_objBetween = new Dictionary<string, Object>();
            SysEntity.MultiLanguage m_MultiLanguage = g_FH.ConvertEntity<SysEntity.MultiLanguage>(p_dicMultiLanguage, ref m_objBetween);
            ArrayList m_alGrid = new ArrayList();
            string m_OrderBy = "";

            try
            {
                string m_SQL = "SELECT [LanguageKey],[ValueTW],[ValueCN],[ValueUS],[ValueJP],[ValueBR] FROM [dbo].[MultiLanguage]    ";
                using (SqlCommand m_cmd = g_DC.GetQueryCommand(p_EmployeeEntity, m_SQL, m_MultiLanguage, m_objBetween))
                {

                    m_TransResult = g_DC.GetDataTableByGrid(p_EmployeeEntity, m_cmd, p_Param, p_PageSize, p_PageIndex);
                    if (m_TransResult.isSuccess)
                    {
                        m_TransResult = g_DC.GridDataCtrl(m_TransResult, p_PageSize, p_Param);
                    }
                }
            }
            catch (Exception ex)
            {
                m_TransResult.LogMessage = ex.Message;
                m_TransResult.isSuccess = false;
            }
            return m_TransResult;
        }

        public SysEntity.TransResult AddMultiLanguage(SysEntity.Employee p_EmployeeEntity, Dictionary<string, Object> p_dicMultiLanguage)
        {
            SysEntity.TransResult m_TransResult = new SysEntity.TransResult();
            try
            {
                m_TransResult = g_DC.InsertIntoByEntity<SysEntity.MultiLanguage>(p_EmployeeEntity, "MultiLanguage", p_dicMultiLanguage);
            }
            catch (Exception ex)
            {
                m_TransResult.LogMessage = ex.Message;
                m_TransResult.isSuccess = false;
            }

            return m_TransResult;

        }

        public SysEntity.TransResult DelMultiLanguage(SysEntity.Employee p_EmployeeEntity, Dictionary<string, Object> p_dicMultiLanguage)
        {
            SysEntity.TransResult m_TransResult = new SysEntity.TransResult();
            Dictionary<string, Object> m_ArrayList = new Dictionary<string, Object>();

            SysEntity.MultiLanguage m_MultiLanguage = g_FH.ConvertEntity<SysEntity.MultiLanguage>(p_dicMultiLanguage, ref m_ArrayList);

            try
            {
                string m_SQL = "Delete [dbo].[MultiLanguage] Where [LanguageKey]=@LanguageKey";
                using (SqlCommand m_cmd = new SqlCommand(m_SQL))
                {
                    m_cmd.Parameters.Add(new SqlParameter("LanguageKey", m_MultiLanguage.LanguageKey));
                    m_TransResult = g_DC.ExecuteNonQuery(p_EmployeeEntity, m_cmd);
                }
            }
            catch (Exception ex)
            {
                m_TransResult.LogMessage = ex.Message;
                m_TransResult.isSuccess = false;
            }

            return m_TransResult;

        }

        public SysEntity.TransResult UpdMultiLanguage(SysEntity.Employee p_EmployeeEntity, Dictionary<string, Object> p_dicMultiLanguage)
        {
            SysEntity.TransResult m_TransResult = new SysEntity.TransResult();
            Dictionary<string, Object> m_ArrayList = new Dictionary<string, Object>();

            SysEntity.MultiLanguage m_MultiLanguage = g_FH.ConvertEntity<SysEntity.MultiLanguage>(p_dicMultiLanguage, ref m_ArrayList);

            try
            {
                string m_SQL = "";
                using (SqlCommand m_cmd = new SqlCommand())
                {
                    m_SQL = "Update [dbo].[MultiLanguage]  set ValueTW =@ValueTW , ValueCN =@ValueCN,ValueJP =@ValueJP,ValueUS =@ValueUS,ValueBR =@ValueBR,ModifyUser=@ModifyUser , ModifyDate=(Convert(varchar(10),Getdate(),111)), ModifyTime=(Convert(varchar(8),Getdate(),108)) ";
                    m_cmd.Parameters.Add(new SqlParameter("ValueTW", m_MultiLanguage.ValueTW));
                    m_cmd.Parameters.Add(new SqlParameter("ValueCN", m_MultiLanguage.ValueCN));
                    m_cmd.Parameters.Add(new SqlParameter("ValueJP", m_MultiLanguage.ValueJP));
                    m_cmd.Parameters.Add(new SqlParameter("ValueUS", m_MultiLanguage.ValueUS));
                    m_cmd.Parameters.Add(new SqlParameter("ValueBR", m_MultiLanguage.ValueBR));

                    m_cmd.Parameters.Add(new SqlParameter("ModifyUser", p_EmployeeEntity.WorkID));

                    if (m_MultiLanguage.LanguageKey != "")
                    {
                        m_SQL += " where [LanguageKey]=@LanguageKey ";
                        m_cmd.Parameters.Add(new SqlParameter("LanguageKey", m_MultiLanguage.LanguageKey));
                        m_cmd.CommandText = m_SQL;
                        m_TransResult = g_DC.ExecuteNonQuery(p_EmployeeEntity, m_cmd);
                    }
                    else
                    {
                        m_TransResult.LogMessage = "LanguageKey 不可為空!!!";
                        m_TransResult.isSuccess = false;
                    }

                }
            }
            catch (Exception ex)
            {
                m_TransResult.LogMessage = ex.Message;
                m_TransResult.isSuccess = false;
            }

            return m_TransResult;

        }

        #endregion

        #region FileUpload

        public SysEntity.TransResult sysFileUpload(SysEntity.Employee p_EmployeeEntity, Business_Logic.Entity.SysEntity.FileFolder p_FileFolder)
        {
            SysEntity.TransResult m_TransResult = new SysEntity.TransResult();
            Dictionary<string, Object> m_ArrayList = new Dictionary<string, Object>();
            try
            {
                var props = p_FileFolder.GetType().GetProperties();
                string m_SQL = "Insert into [dbo].[FileFolder]  ";

                string m_Columns = "";
                string m_Values = "";
                using (SqlCommand m_cmd = new SqlCommand())
                {
                    foreach (var prop in props)
                    {
                        string m_propName = prop.Name;

                        if (prop.GetValue(p_FileFolder, null) != null)
                        {
                            string m_value = prop.GetValue(p_FileFolder, null).ToString(); // against prop.Name
                            if (m_value != "")
                            {
                                m_Columns += (m_Columns == "" ? "(" : ",") + m_propName;
                                m_Values += (m_Values == "" ? "Values(" : ",") + "@" + m_propName;
                                if (m_propName == "FileEntity")
                                {
                                    m_cmd.Parameters.Add("@FileEntity", System.Data.SqlDbType.Binary, p_FileFolder.FileEntity.Length).Value = p_FileFolder.FileEntity;

                                }
                                else
                                {
                                    m_cmd.Parameters.Add(new SqlParameter(m_propName, m_value));

                                }
                            }
                        }
                    }
                    m_SQL += m_Columns + ",CreateUser,ModifyUser )  " + m_Values + ",'" + p_EmployeeEntity.WorkID + "','" + p_EmployeeEntity.WorkID + "' ) ";
                    m_cmd.CommandText = m_SQL;
                    m_TransResult = g_DC.ExecuteNonQuery(p_EmployeeEntity, m_cmd);
                }
            }
            catch (Exception ex)
            {
                m_TransResult.LogMessage = ex.Message;
                m_TransResult.isSuccess = false;
            }
            return m_TransResult;
        }


        public SysEntity.TransResult DelFile(SysEntity.Employee p_EmployeeEntity, Business_Logic.Entity.SysEntity.FileFolder p_FileFolder)
        {
            SysEntity.TransResult m_TransResult = new SysEntity.TransResult();
            try
            {
                string m_SQL = "Delete [dbo].[FileFolder] Where [FileKey]=@FileKey";
                using (SqlCommand m_cmd = new SqlCommand(m_SQL))
                {
                    m_cmd.Parameters.Add(new SqlParameter("FileKey", p_FileFolder.FileKey));
                    m_TransResult = g_DC.ExecuteNonQuery(p_EmployeeEntity, m_cmd);
                }
            }
            catch (Exception ex)
            {
                m_TransResult.LogMessage = ex.Message;
                m_TransResult.isSuccess = false;
            }
            return m_TransResult;
        }

        public SysEntity.TransResult GetFileFolder(SysEntity.Employee p_EmployeeEntity, Business_Logic.Entity.SysEntity.FileFolder p_FileFolder)
        {
            SysEntity.TransResult m_TransResult = new SysEntity.TransResult();
            Dictionary<string, Object> m_objBetween = new Dictionary<string, Object>();

            try
            {
                string m_SQL = "SELECT [FileKey], [DocID],[SiteFormID],[FileName],[FileType],[FileIndex],[FileEntity],[FieldID] FROM [dbo].[FileFolder]  ";
                using (SqlCommand m_cmd = g_DC.GetQueryCommand(p_EmployeeEntity, m_SQL, p_FileFolder, m_objBetween))
                {
                    m_TransResult = g_DC.GetDataTable(p_EmployeeEntity, m_cmd);
                }
            }
            catch (Exception ex)
            {
                m_TransResult.LogMessage = ex.Message;
                m_TransResult.isSuccess = false;
            }
            return m_TransResult;

        }

        public SysEntity.TransResult GetFileEntity(SysEntity.Employee p_EmployeeEntity, Business_Logic.Entity.SysEntity.FileFolder p_FileFolder)
        {
            SysEntity.TransResult m_TransResult = new SysEntity.TransResult();
            Dictionary<string, Object> m_objBetween = new Dictionary<string, Object>();

            List<Dictionary<string, Object>> m_Results = new List<Dictionary<string, Object>>();

            try
            {
                string m_SQL = "SELECT [FileName],[FileType],[FileIndex],[FileEntity] FROM [dbo].[FileFolder]  ";
                using (SqlCommand m_cmd = g_DC.GetQueryCommand(p_EmployeeEntity, m_SQL, p_FileFolder, m_objBetween))
                {
                    m_TransResult = g_DC.GetDataTable(p_EmployeeEntity, m_cmd);
                }
                if (m_TransResult.isSuccess)
                {
                    foreach (DataRow dr in ((DataTable)m_TransResult.ResultEntity).Rows)
                    {
                        Dictionary<string, Object> m_Result = new Dictionary<string, Object>();
                        m_Result["Attachment"] = g_FH.Decompress((Byte[])dr["FileEntity"]);
                        m_Result["FileName"] = dr["FileName"].ToString();
                        m_Results.Add(m_Result);
                    }
                }
                m_TransResult.isSuccess = true;
                m_TransResult.ResultEntity = m_Results;

            }
            catch (Exception ex)
            {
                m_TransResult.LogMessage = ex.Message;
                m_TransResult.isSuccess = false;
            }
            return m_TransResult;
        }


        #endregion


        #region EmployeeGroup Methods


        public SysEntity.TransResult GetGroups(SysEntity.Employee p_EmployeeEntity, Dictionary<string, Object> p_dicGroupID)
        {
            SysEntity.TransResult m_TransResult = new SysEntity.TransResult();
            Dictionary<string, Object> m_objBetween = new Dictionary<string, Object>();
            SysEntity.EmployeeGroup m_EmployeeGroup = g_FH.ConvertEntity<SysEntity.EmployeeGroup>(p_dicGroupID, ref m_objBetween);

            try
            {
                string m_SQL = "  SELECT  GroupDesc, GroupID,OrderIndex FROM dbo.View_GroupID    ";
                using (SqlCommand m_cmd = g_DC.GetQueryCommand(p_EmployeeEntity, m_SQL, m_EmployeeGroup, m_objBetween))
                {
                    m_TransResult = g_DC.GetDataTable(p_EmployeeEntity, m_cmd);
                }
            }
            catch (Exception ex)
            {
                m_TransResult.LogMessage = ex.Message;
                m_TransResult.isSuccess = false;
            }
            return m_TransResult;
        }
        public SysEntity.TransResult GetViewEmployeeGroup(SysEntity.Employee p_EmployeeEntity, Dictionary<string, Object> p_dicViewDetail)
        {
            SysEntity.TransResult m_TransResult = new SysEntity.TransResult();
            Dictionary<string, Object> m_objBetween = new Dictionary<string, Object>();
            SysEntity.EmployeeGroup m_ViewDetail = g_FH.ConvertEntity<SysEntity.EmployeeGroup>(p_dicViewDetail, ref m_objBetween);

            try
            {
                string m_SQL = "SELECT G.[GroupID],P.[GroupDesc],G.[WorkID],E.[DisplayName] CreateUser  FROM dbo.[EmployeeGroup] G " +
                  "   Left Join( " +
                  "   SELECT [ParamID],[GroupDesc],[GroupID],[OrderIndex] FROM [dbo].[View_GroupID] " +
                  "   ) P on G.GroupID=P.GroupID " +
                  "   Left Join( " +
                  "   SELECT WorkID,DisplayName FROM dbo.[EmployeeInfo]  " +
                  "  ) E on G.CreateUser=E.WorkID   ";

                p_dicViewDetail = g_FH.GetKeyByDic("GroupID", p_dicViewDetail);
                m_ViewDetail = g_FH.ConvertEntity<SysEntity.EmployeeGroup>(p_dicViewDetail, ref m_objBetween);
                using (SqlCommand m_cmd = g_DC.GetQueryCommand(p_EmployeeEntity, m_SQL, m_ViewDetail, m_objBetween))
                {
                    m_TransResult = g_DC.GetDataTable(p_EmployeeEntity, m_cmd);
                }
            }
            catch (Exception ex)
            {
                m_TransResult.LogMessage = ex.Message;
                m_TransResult.isSuccess = false;
            }
            return m_TransResult;
        }


        public SysEntity.TransResult GetEmployeeGroups(SysEntity.Employee p_EmployeeEntity, Dictionary<string, Object> p_dicEmployeeGroup)
        {
            SysEntity.TransResult m_TransResult = new SysEntity.TransResult();
            Dictionary<string, Object> m_objBetween = new Dictionary<string, Object>();
            SysEntity.EmployeeGroup m_EmployeeGroup = g_FH.ConvertEntity<SysEntity.EmployeeGroup>(p_dicEmployeeGroup, ref m_objBetween);

            try
            {
                string m_SQL = "SELECT G.*,GD.GroupDesc , E.DisplayName FROM [dbo].[EmployeeGroup] G left join  EmployeeInfo E on G.WorkID=E.WorkID   left join View_GroupID GD on G.GroupID=GD.GroupID      ";
                using (SqlCommand m_cmd = g_DC.GetQueryCommand(p_EmployeeEntity, m_SQL, m_EmployeeGroup, m_objBetween))
                {
                    m_TransResult = g_DC.GetDataTable(p_EmployeeEntity, m_cmd);
                }
            }
            catch (Exception ex)
            {
                m_TransResult.LogMessage = ex.Message;
                m_TransResult.isSuccess = false;
            }
            return m_TransResult;
        }
        public SysEntity.TransResult GetEmployeeGroupGrid(SysEntity.Employee p_EmployeeEntity, Dictionary<string, Object> p_dicEmployeeGroup, string p_PageIndex, string p_PageSize, Dictionary<string, Object> p_Param)
        {
            SysEntity.TransResult m_TransResult = new SysEntity.TransResult();
            Dictionary<string, Object> m_objBetween = new Dictionary<string, Object>();
            SysEntity.EmployeeGroup m_EmployeeGroup = g_FH.ConvertEntity<SysEntity.EmployeeGroup>(p_dicEmployeeGroup, ref m_objBetween);
            ArrayList m_alGrid = new ArrayList();

            try
            {
                string m_SQL = "SELECT G.[GroupID],P.[GroupDesc],G.[WorkID],E.[DisplayName] CreateUser  FROM dbo.[EmployeeGroup] G " +
                    "   Left Join( " +
                    "   SELECT [ParamID],[GroupDesc],[GroupID],[OrderIndex] FROM [dbo].[View_GroupID] " +
                    "   ) P on G.GroupID=P.GroupID " +
                    "   Left Join( " +
                    "   SELECT WorkID,DisplayName FROM dbo.[EmployeeInfo]  " +
                    "  ) E on G.CreateUser=E.WorkID   ";
                using (SqlCommand m_cmd = g_DC.GetQueryCommand(p_EmployeeEntity, m_SQL, m_EmployeeGroup, m_objBetween))
                {

                    m_TransResult = g_DC.GetDataTableByGrid(p_EmployeeEntity, m_cmd, p_Param, p_PageSize, p_PageIndex);
                    if (m_TransResult.isSuccess)
                    {
                        m_TransResult = g_DC.GridDataCtrl(m_TransResult, p_PageSize, p_Param);
                    }
                }
            }
            catch (Exception ex)
            {
                m_TransResult.LogMessage = ex.Message;
                m_TransResult.isSuccess = false;
            }
            return m_TransResult;
        }

        public SysEntity.TransResult AddEmployeeGroup(SysEntity.Employee p_EmployeeEntity, Dictionary<string, Object> p_dicEmployeeGroup)
        {
            SysEntity.TransResult m_TransResult = new SysEntity.TransResult();
            List<SysEntity.EmployeeGroup> m_EmployeeGroups = new List<SysEntity.EmployeeGroup>();
            Dictionary<string, Object> m_objBetween = new Dictionary<string, Object>();
            try
            {
                string m_GroupID = "";
                if (p_dicEmployeeGroup.ContainsKey("GroupID"))
                {
                    m_GroupID = p_dicEmployeeGroup["GroupID"].ToString();
                }
                if (p_dicEmployeeGroup.ContainsKey("Dtls"))
                {
                    Int32 m_RowIndex = 1;
                    foreach (object dic in (object[])p_dicEmployeeGroup["Dtls"])
                    {
                        Dictionary<string, Object> m_dic = (Dictionary<string, Object>)dic;
                        SysEntity.EmployeeGroup m_EmployeeGroup = g_FH.ConvertEntity<SysEntity.EmployeeGroup>((Dictionary<string, Object>)m_dic, ref m_objBetween);
                        m_EmployeeGroup.GroupID = m_GroupID;
                        m_EmployeeGroups.Add(m_EmployeeGroup);
                        m_RowIndex++;
                    }
                }

                string m_SQL = "";
                using (SqlCommand m_cmd = new SqlCommand())
                {
                    Int32 m_RowIndex = 0;
                    foreach (SysEntity.EmployeeGroup m_EmployeeGroup in m_EmployeeGroups)
                    {
                        var props = m_EmployeeGroup.GetType().GetProperties();
                        m_SQL += "Insert into [dbo].[EmployeeGroup]  ";
                        string m_Columns = "";
                        string m_Values = "";

                        foreach (var prop in props)
                        {
                            string m_propName = prop.Name;
                            if (prop.GetValue(m_EmployeeGroup, null) != null)
                            {
                                string m_value = prop.GetValue(m_EmployeeGroup, null).ToString(); // against prop.Name
                                if (m_value != "")
                                {
                                    m_Columns += (m_Columns == "" ? "(" : ",") + m_propName;
                                    m_Values += (m_Values == "" ? "Values(" : ",") + "@" + m_propName + m_RowIndex.ToString();
                                    m_cmd.Parameters.Add(new SqlParameter(m_propName + m_RowIndex.ToString(), m_value.ToString()));
                                }
                            }
                        }
                        m_SQL += m_Columns + ",CreateUser,ModifyUser )  " + m_Values + ",'" + p_EmployeeEntity.WorkID + "','" + p_EmployeeEntity.WorkID + "' ); ";
                        m_RowIndex++;
                    }

                    m_cmd.CommandText += m_SQL;
                    m_TransResult = g_DC.ExecuteNonQuery(p_EmployeeEntity, m_cmd);
                }
            }
            catch (Exception ex)
            {
                m_TransResult.LogMessage = ex.Message;
                m_TransResult.isSuccess = false;
            }

            return m_TransResult;

        }

        public SysEntity.TransResult UpdEmployeeGroup(SysEntity.Employee p_EmployeeEntity, Dictionary<string, Object> p_dicEmployeeGroup)
        {
            SysEntity.TransResult m_TransResult = new SysEntity.TransResult();
            List<SysEntity.EmployeeGroup> m_EmployeeGroups = new List<SysEntity.EmployeeGroup>();
            Dictionary<string, Object> m_objBetween = new Dictionary<string, Object>();
            string m_TempKey = DateTime.Now.ToString("yyyyMMddHHmmssff");


            try
            {
                string m_GroupID = "";
                if (p_dicEmployeeGroup.ContainsKey("GroupID"))
                {
                    m_GroupID = p_dicEmployeeGroup["GroupID"].ToString();
                    using (SqlCommand m_cmd = new SqlCommand())
                    {
                        m_cmd.Parameters.Add(new SqlParameter("GroupID", m_GroupID));
                        m_cmd.CommandText += "Delete dbo.[EmployeeGroup] where GroupID=@GroupID";
                        m_TransResult = g_DC.ExecuteNonQuery(p_EmployeeEntity, m_cmd);
                    }
                }
                if (m_TransResult.isSuccess)
                {
                    if (p_dicEmployeeGroup.ContainsKey("Dtls"))
                    {
                        foreach (object dic in (object[])p_dicEmployeeGroup["Dtls"])
                        {
                            Dictionary<string, Object> m_dic = (Dictionary<string, Object>)dic;
                            if (!(bool)m_dic["DelChk"])
                            {
                                SysEntity.EmployeeGroup m_EmployeeGroup = g_FH.ConvertEntity<SysEntity.EmployeeGroup>((Dictionary<string, Object>)m_dic, ref m_objBetween);
                                m_EmployeeGroup.GroupID = m_GroupID;
                                m_EmployeeGroups.Add(m_EmployeeGroup);
                            }


                          
                        }
                    }
                    if (m_EmployeeGroups.Count != 0)
                    {
                        string m_SQL = "";
                        using (SqlCommand m_cmd = new SqlCommand())
                        {
                            Int32 m_RowIndex = 0;
                            foreach (SysEntity.EmployeeGroup m_EmployeeGroup in m_EmployeeGroups)
                            {
                                var props = m_EmployeeGroup.GetType().GetProperties();
                                m_SQL += "Insert into [dbo].[EmployeeGroup]  ";
                                string m_Columns = "";
                                string m_Values = "";

                                foreach (var prop in props)
                                {
                                    string m_propName = prop.Name;
                                    if (prop.GetValue(m_EmployeeGroup, null) != null)
                                    {
                                        string m_value = prop.GetValue(m_EmployeeGroup, null).ToString(); // against prop.Name
                                        if (m_value != "")
                                        {
                                            m_Columns += (m_Columns == "" ? "(" : ",") + m_propName;
                                            m_Values += (m_Values == "" ? "Values(" : ",") + "@" + m_propName + m_RowIndex.ToString();
                                            m_cmd.Parameters.Add(new SqlParameter(m_propName + m_RowIndex.ToString(), m_value.ToString()));
                                        }
                                    }
                                }
                                m_SQL += m_Columns + ",CreateUser,ModifyUser )  " + m_Values + ",'" + p_EmployeeEntity.WorkID + "','" + p_EmployeeEntity.WorkID + "' ); ";
                                m_RowIndex++;
                            }
                            m_cmd.CommandText += m_SQL;
                            m_TransResult = g_DC.ExecuteNonQuery(p_EmployeeEntity, m_cmd);
                            if (m_TransResult.isSuccess)
                            {
                                m_TransResult.ResultEntity = m_TempKey;
                            }
                        }
                    }
                    

                }
            }
            catch (Exception ex)
            {
                m_TransResult.LogMessage = ex.Message;
                m_TransResult.isSuccess = false;
            }

            return m_TransResult;

        }

        #endregion

        #region ViewDetail Methods

        public SysEntity.TransResult GetViewDetails(SysEntity.Employee p_EmployeeEntity, Dictionary<string, Object> p_dicViewDetail)
        {
            SysEntity.TransResult m_TransResult = new SysEntity.TransResult();
            Dictionary<string, Object> m_objBetween = new Dictionary<string, Object>();
            SysEntity.ViewDetail m_ViewDetail = g_FH.ConvertEntity<SysEntity.ViewDetail>(p_dicViewDetail, ref m_objBetween);

            try
            {
                string m_SQL = "SELECT * FROM [dbo].[ViewDetail]   ";
                using (SqlCommand m_cmd = g_DC.GetQueryCommand(p_EmployeeEntity, m_SQL, m_ViewDetail, m_objBetween))
                {
                    m_TransResult = g_DC.GetDataTable(p_EmployeeEntity, m_cmd);
                }
            }
            catch (Exception ex)
            {
                m_TransResult.LogMessage = ex.Message;
                m_TransResult.isSuccess = false;
            }
            return m_TransResult;
        }
        
        public SysEntity.TransResult GetViewDetailHistory(SysEntity.Employee p_EmployeeEntity, Dictionary<string, Object> p_dicViewDetail)
        {
            SysEntity.TransResult m_TransResult = new SysEntity.TransResult();
            Dictionary<string, Object> m_objBetween = new Dictionary<string, Object>();
            SysEntity.ViewDetail m_ViewDetail = g_FH.ConvertEntity<SysEntity.ViewDetail>(p_dicViewDetail, ref m_objBetween);
            ArrayList m_alGrid = new ArrayList();
            try
            {
                string m_SQL = "select distinct SiteFormID,TempKey,ActionUser,ActionType,convert(varchar, convert(datetime, SUBSTRING(TempKey,1,4)+'/'+SUBSTRING(TempKey,5,2)+'/'+SUBSTRING(TempKey,7,2)+' '+SUBSTRING(TempKey,9,2)+':'+SUBSTRING(TempKey,11,2)+':'+SUBSTRING(TempKey,13,2),120),120) ActionTime from  (select T.TempKey,T.SiteFormID,T.ActionType,ActionUser,S.AgentName from [dbo].[ViewDetail_Temp] T left join [dbo].[SiteForm] S on T.SiteFormID=S.SiteFormID where status='L' ) ViewDetail_Temp  ";
                using (SqlCommand m_cmd = g_DC.GetQueryCommand(p_EmployeeEntity, m_SQL, m_ViewDetail, m_objBetween))
                {
                    m_TransResult = g_DC.GetDataTable(p_EmployeeEntity, m_cmd);
                }
            }
            catch (Exception ex)
            {
                m_TransResult.LogMessage = ex.Message;
                m_TransResult.isSuccess = false;
            }
            return m_TransResult;
        }

        public SysEntity.TransResult GetViewDetailGrid(SysEntity.Employee p_EmployeeEntity, Dictionary<string, Object> p_dicViewDetail, string p_PageIndex, string p_PageSize, Dictionary<string, Object> p_Param)
        {
            SysEntity.TransResult m_TransResult = new SysEntity.TransResult();
            Dictionary<string, Object> m_objBetween = new Dictionary<string, Object>();
            SysEntity.ViewDetail m_ViewDetail = g_FH.ConvertEntity<SysEntity.ViewDetail>(p_dicViewDetail, ref m_objBetween);
            ArrayList m_alGrid = new ArrayList();

            try
            {
                string m_SQL = "SELECT  M.AgentName,D.[SiteFormID],[SectionID] ,[ControlType],[SectionType], [ViewDetailID] ,[FieldID],[QueryID] FROM [dbo].SiteForm M Left Join [dbo].[ViewDetail] D on M.SiteFormID=D.SiteFormID   ";
                using (SqlCommand m_cmd = g_DC.GetQueryCommand(p_EmployeeEntity, m_SQL, m_ViewDetail, m_objBetween))
                {

                    m_TransResult = g_DC.GetDataTableByGrid(p_EmployeeEntity, m_cmd, p_Param, p_PageSize, p_PageIndex);
                    if (m_TransResult.isSuccess)
                    {
                        m_TransResult = g_DC.GridDataCtrl(m_TransResult, p_PageSize, p_Param);
                    }
                }
            }
            catch (Exception ex)
            {
                m_TransResult.LogMessage = ex.Message;
                m_TransResult.isSuccess = false;
            }
            return m_TransResult;
        }

        public SysEntity.TransResult AddViewDetail(SysEntity.Employee p_EmployeeEntity, Dictionary<string, Object> p_dicViewDetail)
        {
            SysEntity.TransResult m_TransResult = new SysEntity.TransResult();
            List<SysEntity.ViewDetail> m_ViewDetails = new List<SysEntity.ViewDetail>();
            Dictionary<string, Object> m_objBetween = new Dictionary<string, Object>();
            try
            {
                string m_SiteFormID = "";
                if (p_dicViewDetail.ContainsKey("SiteFormID"))
                {
                    m_SiteFormID = p_dicViewDetail["SiteFormID"].ToString();
                }
                if (p_dicViewDetail.ContainsKey("Dtls"))
                {
                    Int32 m_RowIndex = 1;
                    foreach (object dic in (object[])p_dicViewDetail["Dtls"])
                    {
                        Dictionary<string, Object> m_dic = (Dictionary<string, Object>)dic;
                        SysEntity.ViewDetail m_ViewDetail = g_FH.ConvertEntity<SysEntity.ViewDetail>((Dictionary<string, Object>)m_dic, ref m_objBetween);
                        m_ViewDetail.ViewDetailID = m_RowIndex.ToString().PadLeft(3, '0');
                        m_ViewDetail.SiteFormID = m_SiteFormID;
                        m_ViewDetails.Add(m_ViewDetail);
                        m_RowIndex++;
                    }
                }

                string m_SQL = "";
                using (SqlCommand m_cmd = new SqlCommand())
                {
                    Int32 m_RowIndex = 0;
                    foreach (SysEntity.ViewDetail m_ViewDetail in m_ViewDetails)
                    {
                        var props = m_ViewDetail.GetType().GetProperties();
                        m_SQL += "Insert into [dbo].[ViewDetail]  ";
                        string m_Columns = "";
                        string m_Values = "";

                        foreach (var prop in props)
                        {
                            string m_propName = prop.Name;
                            if (prop.GetValue(m_ViewDetail, null) != null)
                            {
                                string m_value = prop.GetValue(m_ViewDetail, null).ToString(); // against prop.Name
                                if (m_value != "")
                                {
                                    m_Columns += (m_Columns == "" ? "(" : ",") + m_propName;
                                    m_Values += (m_Values == "" ? "Values(" : ",") + "@" + m_propName + m_RowIndex.ToString();
                                    m_cmd.Parameters.Add(new SqlParameter(m_propName + m_RowIndex.ToString(), m_value.ToString()));
                                }
                            }
                        }
                        m_SQL += m_Columns + ",CreateUser,ModifyUser )  " + m_Values + ",'" + p_EmployeeEntity.WorkID + "','" + p_EmployeeEntity.WorkID + "' ); ";
                        m_RowIndex++;
                    }

                    m_cmd.CommandText += m_SQL;
                    m_TransResult = g_DC.ExecuteNonQuery(p_EmployeeEntity, m_cmd);
                }
            }
            catch (Exception ex)
            {
                m_TransResult.LogMessage = ex.Message;
                m_TransResult.isSuccess = false;
            }

            return m_TransResult;

        }

        public SysEntity.TransResult DelViewDetail(SysEntity.Employee p_EmployeeEntity, Dictionary<string, Object> p_dicViewDetail)
        {
            SysEntity.TransResult m_TransResult = new SysEntity.TransResult();
            Dictionary<string, Object> m_ArrayList = new Dictionary<string, Object>();

            SysEntity.ViewDetail m_ViewDetail = g_FH.ConvertEntity<SysEntity.ViewDetail>(p_dicViewDetail, ref m_ArrayList);

            try
            {
                string m_SQL = "Delete [dbo].[ViewDetail] Where [SiteFormID]=@SiteFormID";
                using (SqlCommand m_cmd = new SqlCommand(m_SQL))
                {
                    m_cmd.Parameters.Add(new SqlParameter("ViewDetailID", m_ViewDetail.SiteFormID));
                    m_TransResult = g_DC.ExecuteNonQuery(p_EmployeeEntity, m_cmd);
                }
            }
            catch (Exception ex)
            {
                m_TransResult.LogMessage = ex.Message;
                m_TransResult.isSuccess = false;
            }

            return m_TransResult;

        }

        public SysEntity.TransResult UpdViewDetail(SysEntity.Employee p_EmployeeEntity, Dictionary<string, Object> p_dicViewDetail)
        {
            SysEntity.TransResult m_TransResult = new SysEntity.TransResult();
            List<SysEntity.ViewDetail> m_ViewDetails = new List<SysEntity.ViewDetail>();
            Dictionary<string, Object> m_objBetween = new Dictionary<string, Object>();
            string m_TempKey = DateTime.Now.ToString("yyyyMMddHHmmssff");


            try
            {
                string m_SiteFormID = "";
                if (p_dicViewDetail.ContainsKey("SiteFormID"))
                {
                    m_SiteFormID = p_dicViewDetail["SiteFormID"].ToString();
                    using (SqlCommand m_cmd = new SqlCommand())
                    {
                        m_cmd.Parameters.Add(new SqlParameter("SiteFormID", m_SiteFormID));
                        m_cmd.Parameters.Add(new SqlParameter("TempKey", m_TempKey));
                        m_cmd.CommandText += "INSERT INTO dbo.[ViewDetail_Temp] SELECT @TempKey TempKey ,V.*,'L' Status,'Modify'ActionType,'" + p_EmployeeEntity.WorkID + "' ActionUser from  [dbo].[ViewDetail] V where V.SiteFormID=@SiteFormID";
                        m_TransResult = g_DC.ExecuteNonQuery(p_EmployeeEntity, m_cmd);
                    }
                    if (m_TransResult.isSuccess)
                    {
                        using (SqlCommand m_cmd = new SqlCommand())
                        {
                            m_cmd.Parameters.Add(new SqlParameter("SiteFormID", m_SiteFormID));
                            m_cmd.CommandText += "Delete dbo.[ViewDetail] where SiteFormID=@SiteFormID";
                            m_TransResult = g_DC.ExecuteNonQuery(p_EmployeeEntity, m_cmd);
                        }
                    }
                }
                if (m_TransResult.isSuccess)
                {
                    if (p_dicViewDetail.ContainsKey("Dtls"))
                    {
                        Int32 m_RowIndex = 1;
                        foreach (object dic in (object[])p_dicViewDetail["Dtls"])
                        {
                            Dictionary<string, Object> m_dic = (Dictionary<string, Object>)dic;
                            SysEntity.ViewDetail m_ViewDetail = g_FH.ConvertEntity<SysEntity.ViewDetail>((Dictionary<string, Object>)m_dic, ref m_objBetween);
                            m_ViewDetail.ViewDetailID = m_RowIndex.ToString().PadLeft(3, '0');
                            m_ViewDetail.SiteFormID = m_SiteFormID;
                            m_ViewDetails.Add(m_ViewDetail);
                            m_RowIndex++;
                        }
                    }

                    string m_SQL = "";
                    using (SqlCommand m_cmd = new SqlCommand())
                    {
                        Int32 m_RowIndex = 0;
                        foreach (SysEntity.ViewDetail m_ViewDetail in m_ViewDetails)
                        {
                            var props = m_ViewDetail.GetType().GetProperties();
                            m_SQL += "Insert into [dbo].[ViewDetail]  ";
                            string m_Columns = "";
                            string m_Values = "";

                            foreach (var prop in props)
                            {
                                string m_propName = prop.Name;
                                if (prop.GetValue(m_ViewDetail, null) != null)
                                {
                                    string m_value = prop.GetValue(m_ViewDetail, null).ToString(); // against prop.Name
                                    if (m_value != "")
                                    {
                                        m_Columns += (m_Columns == "" ? "(" : ",") + m_propName;
                                        m_Values += (m_Values == "" ? "Values(" : ",") + "@" + m_propName + m_RowIndex.ToString();
                                        m_cmd.Parameters.Add(new SqlParameter(m_propName + m_RowIndex.ToString(), m_value.ToString()));
                                    }
                                }
                            }
                            m_SQL += m_Columns + ",CreateUser,ModifyUser )  " + m_Values + ",'" + p_EmployeeEntity.WorkID + "','" + p_EmployeeEntity.WorkID + "' ); ";
                            m_RowIndex++;
                        }
                        m_cmd.CommandText += m_SQL;
                        m_TransResult = g_DC.ExecuteNonQuery(p_EmployeeEntity, m_cmd);
                        if (m_TransResult.isSuccess)
                        {
                            m_TransResult.ResultEntity = m_TempKey;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                m_TransResult.LogMessage = ex.Message;
                m_TransResult.isSuccess = false;
            }

            return m_TransResult;

        }

        public SysEntity.TransResult GetViewMaintain(SysEntity.Employee p_EmployeeEntity, Dictionary<string, Object> p_dicViewDetail)
        {
            SysEntity.TransResult m_TransResult = new SysEntity.TransResult();
            Dictionary<string, Object> m_objBetween = new Dictionary<string, Object>();
            SysEntity.ViewDetail m_ViewDetail = g_FH.ConvertEntity<SysEntity.ViewDetail>(p_dicViewDetail, ref m_objBetween);

            try
            {
                string m_SQL = "SELECT   M.AgentName,M.SiteID,M.FormID,D.*  FROM [dbo].[SiteForm] M , (select D.*,V.Source ValidSource ,V.Parameter ValidParam from [dbo].[ViewDetail] D left join  [dbo].[ValidatedInfo] V on D.ValidID=V.ValidID ) D left join [dbo].[EventSetting] E on D.EventID=E.EventID  where M.SiteFormID=D.SiteFormID  ";
                p_dicViewDetail = g_FH.GetKeyByDic("SiteFormID", p_dicViewDetail);
                m_ViewDetail = g_FH.ConvertEntity<SysEntity.ViewDetail>(p_dicViewDetail, ref m_objBetween);
                using (SqlCommand m_cmd = g_DC.GetQueryCommand(p_EmployeeEntity, m_SQL, m_ViewDetail, m_objBetween))
                {
                    m_TransResult = g_DC.GetDataTable(p_EmployeeEntity, m_cmd);
                }
            }
            catch (Exception ex)
            {
                m_TransResult.LogMessage = ex.Message;
                m_TransResult.isSuccess = false;
            }
            return m_TransResult;
        }

        public SysEntity.TransResult AddViewDetail_Temp(SysEntity.Employee p_EmployeeEntity, Dictionary<string, Object> p_dicViewDetail)
        {
            SysEntity.TransResult m_TransResult = new SysEntity.TransResult();
            List<SysEntity.ViewDetail> m_ViewDetails = new List<SysEntity.ViewDetail>();
            Dictionary<string, Object> m_objBetween = new Dictionary<string, Object>();
            string m_TempKey = DateTime.Now.ToString("yyyyMMddHHmmssff");
            try
            {
                string m_SiteFormID = "";
                if (p_dicViewDetail.ContainsKey("SiteFormID"))
                {
                    m_SiteFormID = p_dicViewDetail["SiteFormID"].ToString();
                }
                if (p_dicViewDetail.ContainsKey("Dtls"))
                {
                    Int32 m_RowIndex = 1;
                    foreach (object dic in (object[])p_dicViewDetail["Dtls"])
                    {
                        Dictionary<string, Object> m_dic = (Dictionary<string, Object>)dic;
                        SysEntity.ViewDetail m_ViewDetail = g_FH.ConvertEntity<SysEntity.ViewDetail>((Dictionary<string, Object>)m_dic, ref m_objBetween);
                        m_ViewDetail.ViewDetailID = m_RowIndex.ToString().PadLeft(3, '0');
                        m_ViewDetail.SiteFormID = m_SiteFormID;
                        m_ViewDetails.Add(m_ViewDetail);
                        m_RowIndex++;
                    }
                }

                string m_SQL = "";
                using (SqlCommand m_cmd = new SqlCommand())
                {
                    Int32 m_RowIndex = 0;
                    foreach (SysEntity.ViewDetail m_ViewDetail in m_ViewDetails)
                    {
                        var props = m_ViewDetail.GetType().GetProperties();
                        m_SQL += "Insert into [dbo].[ViewDetail_Temp]  ";
                        string m_Columns = "";
                        string m_Values = "";

                        foreach (var prop in props)
                        {
                            string m_propName = prop.Name;
                            if (prop.GetValue(m_ViewDetail, null) != null)
                            {
                                string m_value = prop.GetValue(m_ViewDetail, null).ToString(); // against prop.Name
                                if (m_value != "")
                                {
                                    m_Columns += (m_Columns == "" ? "(" : ",") + m_propName;
                                    m_Values += (m_Values == "" ? "Values(" : ",") + "@" + m_propName + m_RowIndex.ToString();
                                    m_cmd.Parameters.Add(new SqlParameter(m_propName + m_RowIndex.ToString(), m_value.ToString()));
                                }
                            }
                        }
                        m_SQL += m_Columns + ",TempKey,CreateUser,ModifyUser,Status )  " + m_Values + ",'" + m_TempKey + "','" + p_EmployeeEntity.WorkID + "','" + p_EmployeeEntity.WorkID + "','P' ); ";
                        m_RowIndex++;
                    }

                    m_cmd.CommandText += m_SQL;
                    m_TransResult = g_DC.ExecuteNonQuery(p_EmployeeEntity, m_cmd);

                    if (m_TransResult.isSuccess)
                    {
                        m_TransResult.ResultEntity = m_TempKey;
                    }

                }
            }
            catch (Exception ex)
            {
                m_TransResult.LogMessage = ex.Message;
                m_TransResult.isSuccess = false;
            }

            return m_TransResult;

        }

        public SysEntity.TransResult RollbackViewDetail(SysEntity.Employee p_EmployeeEntity, Dictionary<string, Object> p_dicViewDetail)
        {
            SysEntity.TransResult m_TransResult = new SysEntity.TransResult();
            List<SysEntity.ViewDetail> m_ViewDetails = new List<SysEntity.ViewDetail>();
            Dictionary<string, Object> m_objBetween = new Dictionary<string, Object>();
            string m_TempKey = DateTime.Now.ToString("yyyyMMddHHmmssff");

            try
            {
                string m_SiteFormID = "";
                if (p_dicViewDetail.ContainsKey("SiteFormID"))
                {
                    m_SiteFormID = p_dicViewDetail["SiteFormID"].ToString();
                }
                string m_OldTempKey = "";
                if (p_dicViewDetail.ContainsKey("OldTempKey"))
                {
                    m_OldTempKey = p_dicViewDetail["OldTempKey"].ToString();
                    using (SqlCommand m_cmd = new SqlCommand())
                    {
                        m_cmd.Parameters.Add(new SqlParameter("SiteFormID", m_SiteFormID));
                        m_cmd.Parameters.Add(new SqlParameter("TempKey", m_TempKey));

                        m_cmd.CommandText += "INSERT INTO dbo.[ViewDetail_Temp] SELECT @TempKey TempKey ,V.*,'L' Status,'Rollback'ActionType,'" + p_EmployeeEntity.WorkID + "' ActionUser from  [dbo].[ViewDetail] V where V.SiteFormID=@SiteFormID";
                        m_TransResult = g_DC.ExecuteNonQuery(p_EmployeeEntity, m_cmd);
                    }
                    if (m_TransResult.isSuccess)
                    {
                        using (SqlCommand m_cmd = new SqlCommand())
                        {
                            m_cmd.Parameters.Add(new SqlParameter("SiteFormID", m_SiteFormID));
                            m_cmd.CommandText += "Delete dbo.[ViewDetail] where SiteFormID=@SiteFormID";
                            m_TransResult = g_DC.ExecuteNonQuery(p_EmployeeEntity, m_cmd);
                        }
                    }
                    if (m_TransResult.isSuccess)
                    {
                        using (SqlCommand m_cmd = new SqlCommand())
                        {
                            m_cmd.Parameters.Add(new SqlParameter("OldTempKey", m_OldTempKey));
                            m_cmd.CommandText = "INSERT INTO dbo.[ViewDetail] SELECT [SiteFormID],[ViewDetailID] ,[SectionID],[SectionType],[ControlType],[FieldID],[QueryID],[ControlRef],[ValidID],[DefaultValue],[Height],[Width],[EventID],[MaxLength],[ModifyKey],[QueryMethod], (CONVERT([varchar](10),getdate(),(111))),(CONVERT([varchar](8),getdate(),(108))),'" + p_EmployeeEntity.WorkID + "' [CreateUser],'" + p_EmployeeEntity.WorkID + "' [ModifyUser]  ,(CONVERT([varchar](10),getdate(),(111))),(CONVERT([varchar](8),getdate(),(108)))  FROM [dbo].[ViewDetail_Temp] V where V.TempKey=@OldTempKey";
                            m_TransResult = g_DC.ExecuteNonQuery(p_EmployeeEntity, m_cmd);
                        }
                    }
                }

            }
            catch (Exception ex)
            {
                m_TransResult.LogMessage = ex.Message;
                m_TransResult.isSuccess = false;
            }

            return m_TransResult;

        }

        public DataTable GetMaintainButton(SysEntity.Employee p_EmployeeEntity, string p_SiteFormID)
        {
            DataTable m_dt = new DataTable();
            string m_SQL = "SELECT * from [dbo].[ViewDetail]  where SiteFormID=@SiteFormID and [SectionID]='divButton' and [SectionType]='Maintain' ";

            using (SqlCommand m_cmd = new SqlCommand(m_SQL))
            {
                m_cmd.Parameters.Add(new SqlParameter("SiteFormID", p_SiteFormID));
                SysEntity.TransResult m_TransResult = new SysEntity.TransResult();
                m_cmd.CommandText += " order by [ViewDetailID]  ";
                m_TransResult = g_DC.GetDataTable(p_EmployeeEntity, m_cmd);
                if (m_TransResult.isSuccess)
                {
                    m_dt = (DataTable)m_TransResult.ResultEntity;
                }
            }
            return m_dt;
        }

        public DataTable GetViewDetailBySiteFormID(SysEntity.Employee p_EmployeeEntity, string p_SiteFormID, string p_TempKey = "")
        {
            DataTable m_dt = new DataTable();

            string m_SQL = "SELECT DISTINCT M.AgentName,M.SiteDesc,D.SectionID  FROM (select D.Description SiteDesc,M.* from  [dbo].[SiteForm] M Left join [dbo].[Site] D on M.SiteID=D.SiteID ) M" +
                " , (select D.*,V.Source ValidSource ,V.Parameter ValidParam from [dbo].[ViewDetail] D left join  [dbo].[ValidatedInfo] V on D.ValidID=V.ValidID ) D left join [dbo].[EventSetting] E" +
                " on D.EventID=E.EventID  where M.SiteFormID=D.SiteFormID  AND M.SiteFormID=@SiteFormID AND D.SectionID <> 'divMaintain'  AND D.SectionType <> 'Maintain'  ";


            using (SqlCommand m_cmd = new SqlCommand(m_SQL))
            {
                m_cmd.Parameters.Add(new SqlParameter("SiteFormID", p_SiteFormID));
                if (p_TempKey != "")
                {
                    m_cmd.CommandText = m_cmd.CommandText.Replace("[dbo].[ViewDetail]", "[dbo].[ViewDetail_Temp]");
                    m_cmd.CommandText += " AND TempKey=@TempKey";
                    m_cmd.Parameters.Add(new SqlParameter("TempKey", p_TempKey));
                }

                SysEntity.TransResult m_TransResult = new SysEntity.TransResult();

                m_TransResult = g_DC.GetDataTable(p_EmployeeEntity, m_cmd);
                if (m_TransResult.isSuccess)
                {
                    m_dt = (DataTable)m_TransResult.ResultEntity;
                }
            }
            return m_dt;
        }

        public DataTable GetViewDetail(SysEntity.Employee p_EmployeeEntity, string p_SiteFormID, string p_SectionID, string p_SectionType = "", string p_TempKey = "")
        {

            DataTable m_dt = new DataTable();
            string m_SQL = "SELECT M.SiteFormID,[Admin],[NotesPath],[AgentName],M.[ValidID] as SubmitValidID,[DocLockTime] ,D.[ViewDetailID] ,[SectionID] ,[SectionType],[ControlType],[FieldID],[QueryID],[ControlRef],D.[ValidID],[DefaultValue],[Height],[Width],D.[EventID],[MaxLength],[ModifyKey],[QueryMethod] ";
            m_SQL += ",ValidSource,ValidParam,E.[EventModel],[EventAction],[EventDescription],[EventActionType],[EventFollow],[EventRef],[EventCtrl]  FROM [dbo].[SiteForm] M , (select  [SiteFormID] ,[ViewDetailID] ,[SectionID] ,[SectionType],[ControlType],[FieldID],[QueryID],[ControlRef],D.[ValidID],[DefaultValue],[Height],[Width],[EventID],[MaxLength],[ModifyKey],[QueryMethod] ";
            
            m_SQL += " ,V.Source ValidSource ,V.Parameter ValidParam from [dbo].[ViewDetail] D left join  [dbo].[ValidatedInfo] V on D.ValidID=V.ValidID ) D left join ";
            m_SQL += " ( SELECT [EventID],[EventModel],[EventAction],[EventDescription],[EventActionType],[EventFollow],[EventRef],[EventCtrl] FROM [dbo].[EventSetting] ) E";
            m_SQL += " on D.EventID=E.EventID  where M.SiteFormID=D.SiteFormID  AND M.SiteFormID=@SiteFormID AND D.SectionID=@SectionID ";
            
            
            if (p_SectionType != "")
            {
                m_SQL += " AND D.SectionType=@SectionType ";
            }
            using (SqlCommand m_cmd = new SqlCommand(m_SQL))
            {

                if (p_TempKey != "")
                {
                    m_cmd.CommandText = m_cmd.CommandText.Replace("[dbo].[ViewDetail]", "[dbo].[ViewDetail_Temp]");
                    m_cmd.CommandText += " AND TempKey=@TempKey ";
                    m_cmd.Parameters.Add(new SqlParameter("TempKey", p_TempKey));
                }

                m_cmd.Parameters.Add(new SqlParameter("SiteFormID", p_SiteFormID));
                m_cmd.Parameters.Add(new SqlParameter("SectionID", p_SectionID));
                if (p_SectionType != "")
                {
                    m_cmd.Parameters.Add(new SqlParameter("SectionType", p_SectionType));

                }

                SysEntity.TransResult m_TransResult = new SysEntity.TransResult();

                m_cmd.CommandText += " order by (CASE SectionType WHEN 'Detail' THEN 'Z'+[ViewDetailID] else [ViewDetailID] end) , [SectionType] desc  ";

                m_TransResult = g_DC.GetDataTable(p_EmployeeEntity, m_cmd);
                if (m_TransResult.isSuccess)
                {
                    m_dt = (DataTable)m_TransResult.ResultEntity;
                }
            }
            return m_dt;
        }

        public DataTable GetViewDetailItems(SysEntity.Employee p_EmployeeEntity, string p_SiteFormID, string p_SectionID, string p_ItemKey)
        {

            DataTable m_dt = new DataTable();
            string m_SQL = "SELECT M.SiteFormID,[Admin],[NotesPath],[AgentName],M.[ValidID] as SubmitValidID,[DocLockTime] ,D.[ViewDetailID] ,[SectionID] ,[SectionType],[ControlType],[FieldID],[QueryID],[ControlRef],D.[ValidID],[DefaultValue],[Height],[Width],D.[EventID],[MaxLength],[ModifyKey],[QueryMethod] ";
            m_SQL += ",ValidSource,ValidParam,E.[EventModel],[EventAction],[EventDescription],[EventActionType],[EventFollow],[EventRef],[EventCtrl]  FROM [dbo].[SiteForm] M , (select  [SiteFormID] ,[ViewDetailID] ,[SectionID] ,[SectionType],[ControlType],[FieldID],[QueryID],[ControlRef],D.[ValidID],[DefaultValue],[Height],[Width],[EventID],[MaxLength],[ModifyKey],[QueryMethod] ";

            m_SQL += " ,V.Source ValidSource ,V.Parameter ValidParam from (select * from [dbo].[ViewDetailItems] where ItemKey=@ItemKey ) D left join  [dbo].[ValidatedInfo] V on D.ValidID=V.ValidID ) D left join ";
            m_SQL += " ( SELECT [EventID],[EventModel],[EventAction],[EventDescription],[EventActionType],[EventFollow],[EventRef],[EventCtrl] FROM [dbo].[EventSetting] ) E";
            m_SQL += " on D.EventID=E.EventID  where M.SiteFormID=D.SiteFormID  AND M.SiteFormID=@SiteFormID AND D.SectionID=@SectionID ";


            
            using (SqlCommand m_cmd = new SqlCommand(m_SQL))
            {

               

                m_cmd.Parameters.Add(new SqlParameter("SiteFormID", p_SiteFormID));
                m_cmd.Parameters.Add(new SqlParameter("SectionID", p_SectionID));
                m_cmd.Parameters.Add(new SqlParameter("ItemKey", p_ItemKey));

                SysEntity.TransResult m_TransResult = new SysEntity.TransResult();

                m_cmd.CommandText += " order by (CASE SectionType WHEN 'Detail' THEN 'Z'+[ViewDetailID] else [ViewDetailID] end) , [SectionType] desc  ";

                m_TransResult = g_DC.GetDataTable(p_EmployeeEntity, m_cmd);
                if (m_TransResult.isSuccess)
                {
                    m_dt = (DataTable)m_TransResult.ResultEntity;
                }
            }
            return m_dt;
        }

        public DataTable GetViewDetailItemValues(SysEntity.Employee p_EmployeeEntity, string p_SiteFormID, string p_SectionID, string p_ItemKey)
        {

            DataTable m_dt = new DataTable();
            string m_SQL = "SELECT D.[ViewDetailID] ,[SectionID] ,[SectionType],[ControlType],[FieldID],[QueryID],[ControlRef],[ValidID],[DefaultValue],[Height],[Width],[EventID],[MaxLength],[ModifyKey],[QueryMethod]  ";

            m_SQL += "  FROM  [dbo].[ViewDetailItems] D where  D.SiteFormID=@SiteFormID and ItemKey=@ItemKey   ";

            using (SqlCommand m_cmd = new SqlCommand(m_SQL))
            {
                m_cmd.Parameters.Add(new SqlParameter("SiteFormID", p_SiteFormID));
                m_cmd.Parameters.Add(new SqlParameter("SectionID", p_SectionID));
                m_cmd.Parameters.Add(new SqlParameter("ItemKey", p_ItemKey));

                SysEntity.TransResult m_TransResult = new SysEntity.TransResult();

                m_cmd.CommandText += " order by ViewDetailID sec  ";

                m_TransResult = g_DC.GetDataTable(p_EmployeeEntity, m_cmd);
                if (m_TransResult.isSuccess)
                {
                    m_dt = (DataTable)m_TransResult.ResultEntity;
                }
            }
            return m_dt;
        }


        public DataTable GetViewAuthByDocID(SysEntity.Employee p_EmployeeEntity, string p_DocID, string p_SiteFormID, string p_AuthType)
        {

            DataTable m_dt = new DataTable();
            string m_SQL = "";
            m_SQL = "SELECT  [AuthID],[ViewDetailID],[Visible],[Edit],[Style] FROM  [dbo].[Doc] D left join  [dbo].[ViewAuth] V  on D.SiteFormID=V.SiteFormID and D.Level=V.Level where D.[DocID]=@DocID and [AuthID]='ALL' union ALL ";

            m_SQL += "SELECT  [AuthID],[ViewDetailID],[Visible],[Edit],[Style] FROM  [dbo].[Doc] D left join  [dbo].[ViewAuth] V  on D.SiteFormID=V.SiteFormID and D.Level=V.Level where D.[DocID]=@DocID and [AuthID]<>'ALL'";

            if (p_AuthType == "Initial")
            {
                m_SQL = "  SELECT  [AuthID],[ViewDetailID],[Visible],[Edit],[Style] FROM  [dbo].[ViewAuth] D where D.SiteFormID=@SiteFormID and D.Level='N' and [AuthID]='ALL'  union ALL  ";
                m_SQL = "  SELECT  [AuthID],[ViewDetailID],[Visible],[Edit],[Style] FROM  [dbo].[ViewAuth] D where D.SiteFormID=@SiteFormID and D.Level='N' and [AuthID] <> 'ALL'  ";

                using (SqlCommand m_cmd = new SqlCommand(m_SQL))
                {
                    m_cmd.Parameters.Add(new SqlParameter("SiteFormID", p_SiteFormID));
                    SysEntity.TransResult m_TransResult = new SysEntity.TransResult();
                    m_TransResult = g_DC.GetDataTable(p_EmployeeEntity, m_cmd);
                    if (m_TransResult.isSuccess)
                    {
                        m_dt = (DataTable)m_TransResult.ResultEntity;
                    }
                }
            }
            else
            {
                using (SqlCommand m_cmd = new SqlCommand(m_SQL))
                {
                    m_cmd.Parameters.Add(new SqlParameter("DocID", p_DocID));

                    SysEntity.TransResult m_TransResult = new SysEntity.TransResult();


                    m_TransResult = g_DC.GetDataTable(p_EmployeeEntity, m_cmd);
                    if (m_TransResult.isSuccess)
                    {
                        m_dt = (DataTable)m_TransResult.ResultEntity;
                    }
                }
            }


            return m_dt;
        }

        public DataTable GetViewAuthByAuthKey(SysEntity.Employee p_EmployeeEntity, string p_SiteFormID, string p_AuthType)
        {

            DataTable m_dt = new DataTable();
            string m_SQL = "";
            m_SQL = " SELECT  [AuthID],[ViewDetailID],[Visible],[Edit],[Style],[AuthRule],[AuthRuleValue] FROM  [dbo].[ViewAuth] V where V.SiteFormID=@SiteFormID and V.AuthType='Maintain' and [AuthID]='ALL' union ALL ";
            m_SQL += "SELECT  [AuthID],[ViewDetailID],[Visible],[Edit],[Style],[AuthRule],[AuthRuleValue] FROM  [dbo].[ViewAuth] V where V.SiteFormID=@SiteFormID and V.AuthType='Maintain' and [AuthID]<>'ALL'";

            if (p_AuthType == "Initial")
            {
                m_SQL = "  SELECT  [AuthID],[ViewDetailID],[Visible],[Edit],[Style] FROM  [dbo].[ViewAuth] V where V.SiteFormID=@SiteFormID and V.AuthType='Initial'  and V.[AuthID]='ALL'  union ALL  ";
                m_SQL += " SELECT  [AuthID],[ViewDetailID],[Visible],[Edit],[Style] FROM  [dbo].[ViewAuth] V where V.SiteFormID=@SiteFormID and V.AuthType='Initial'  and V.[AuthID] <> 'ALL'  ";

                using (SqlCommand m_cmd = new SqlCommand(m_SQL))
                {
                    m_cmd.Parameters.Add(new SqlParameter("SiteFormID", p_SiteFormID));
                    SysEntity.TransResult m_TransResult = new SysEntity.TransResult();
                    m_TransResult = g_DC.GetDataTable(p_EmployeeEntity, m_cmd);
                    if (m_TransResult.isSuccess)
                    {
                        m_dt = (DataTable)m_TransResult.ResultEntity;
                    }
                }
            }
            else
            {
                using (SqlCommand m_cmd = new SqlCommand(m_SQL))
                {
                    m_cmd.Parameters.Add(new SqlParameter("SiteFormID", p_SiteFormID));
                    SysEntity.TransResult m_TransResult = new SysEntity.TransResult();
                    m_TransResult = g_DC.GetDataTable(p_EmployeeEntity, m_cmd);
                    if (m_TransResult.isSuccess)
                    {
                        m_dt = (DataTable)m_TransResult.ResultEntity;
                    }
                }
            }


            return m_dt;
        }

        public SysEntity.TransResult GetViewAuthSetting(SysEntity.Employee p_EmployeeEntity, string p_SiteFormID)
        {
            SysEntity.TransResult m_TransResult = new SysEntity.TransResult();
            DataTable m_dt = new DataTable();
            string m_SQL = "";
            m_SQL = "SELECT [SiteFormID],[AuthTable],[AuthKey],[AuthRule] FROM  [dbo].[ViewAuthSetting] where SiteFormID=@SiteFormID ";
            using (SqlCommand m_cmd = new SqlCommand(m_SQL))
            {
                m_cmd.Parameters.Add(new SqlParameter("SiteFormID", p_SiteFormID));
                m_TransResult = g_DC.GetDataTable(p_EmployeeEntity, m_cmd);
                if (m_TransResult.isSuccess)
                {
                    m_dt = (DataTable)m_TransResult.ResultEntity;
                }
            }
            return m_TransResult;
        }

        public DataTable GetViewAuth(SysEntity.Employee p_EmployeeEntity, string p_DocID, string p_SiteFormID, string p_PageStatus)
        {

            DataTable m_dt = new DataTable();
            string m_SQL = "";
            m_SQL = "SELECT  [AuthID],[ViewDetailID],[Visible],[Edit],[Style] FROM  [dbo].[Doc] D left join  [dbo].[ViewAuth] V  on D.SiteFormID=V.SiteFormID and D.Level=V.Level where  D.[DocID]=@DocID and [AuthID]='ALL' union ALL ";

            m_SQL += "SELECT  [AuthID],[ViewDetailID],[Visible],[Edit],[Style] FROM  [dbo].[Doc] D left join  [dbo].[ViewAuth] V  on D.SiteFormID=V.SiteFormID and D.Level=V.Level where  D.[DocID]=@DocID and [AuthID]<>'ALL'";

            if (p_PageStatus == "Initial")
            {
                m_SQL = "  SELECT  [AuthID],[ViewDetailID],[Visible],[Edit],[Style] FROM  [dbo].[ViewAuth] D where D.SiteFormID=@SiteFormID and D.Level='N'  and [AuthID]='ALL'  union ALL  ";
                m_SQL = "  SELECT  [AuthID],[ViewDetailID],[Visible],[Edit],[Style] FROM  [dbo].[ViewAuth] D where D.SiteFormID=@SiteFormID and D.Level='N'  and [AuthID] <> 'ALL'  ";

                using (SqlCommand m_cmd = new SqlCommand(m_SQL))
                {
                    m_cmd.Parameters.Add(new SqlParameter("SiteFormID", p_SiteFormID));
                    SysEntity.TransResult m_TransResult = new SysEntity.TransResult();
                    m_TransResult = g_DC.GetDataTable(p_EmployeeEntity, m_cmd);
                    if (m_TransResult.isSuccess)
                    {
                        m_dt = (DataTable)m_TransResult.ResultEntity;
                    }
                }
            }
            else
            {
                using (SqlCommand m_cmd = new SqlCommand(m_SQL))
                {
                    m_cmd.Parameters.Add(new SqlParameter("DocID", p_DocID));

                    SysEntity.TransResult m_TransResult = new SysEntity.TransResult();


                    m_TransResult = g_DC.GetDataTable(p_EmployeeEntity, m_cmd);
                    if (m_TransResult.isSuccess)
                    {
                        m_dt = (DataTable)m_TransResult.ResultEntity;
                    }
                }
            }


            return m_dt;
        }




        #endregion

        #region EventSetting Methods


        public SysEntity.TransResult GetEventSettings(SysEntity.Employee p_EmployeeEntity, Dictionary<string, Object> p_dicEventSetting)
        {
            SysEntity.TransResult m_TransResult = new SysEntity.TransResult();
            Dictionary<string, Object> m_objBetween = new Dictionary<string, Object>();
            SysEntity.EventSetting m_EventSetting = g_FH.ConvertEntity<SysEntity.EventSetting>(p_dicEventSetting, ref m_objBetween);

            try
            {
                string m_SQL = "SELECT * FROM [dbo].[EventSetting]   ";
                using (SqlCommand m_cmd = g_DC.GetQueryCommand(p_EmployeeEntity, m_SQL, m_EventSetting, m_objBetween))
                {
                    m_TransResult = g_DC.GetDataTable(p_EmployeeEntity, m_cmd);
                }
            }
            catch (Exception ex)
            {
                m_TransResult.LogMessage = ex.Message;
                m_TransResult.isSuccess = false;
            }
            return m_TransResult;
        }

        public SysEntity.TransResult GetEventSettingGrid(SysEntity.Employee p_EmployeeEntity, Dictionary<string, Object> p_dicEventSetting, string p_PageIndex, string p_PageSize, Dictionary<string, Object> p_Param)
        {
            SysEntity.TransResult m_TransResult = new SysEntity.TransResult();
            Dictionary<string, Object> m_objBetween = new Dictionary<string, Object>();
            SysEntity.EventSetting m_EventSetting = g_FH.ConvertEntity<SysEntity.EventSetting>(p_dicEventSetting, ref m_objBetween);
            ArrayList m_alGrid = new ArrayList();


            try
            {
                string m_SQL = "SELECT [EventID],[EventModel],[EventAction],[EventDescription] FROM [dbo].[EventSetting]   ";
                using (SqlCommand m_cmd = g_DC.GetQueryCommand(p_EmployeeEntity, m_SQL, m_EventSetting, m_objBetween))
                {

                    m_TransResult = g_DC.GetDataTableByGrid(p_EmployeeEntity, m_cmd, p_Param, p_PageSize, p_PageIndex);
                    if (m_TransResult.isSuccess)
                    {
                        m_TransResult = g_DC.GridDataCtrl(m_TransResult, p_PageSize, p_Param);
                    }
                }
            }
            catch (Exception ex)
            {
                m_TransResult.LogMessage = ex.Message;
                m_TransResult.isSuccess = false;
            }
            return m_TransResult;
        }

        public SysEntity.TransResult AddEventSetting(SysEntity.Employee p_EmployeeEntity, Dictionary<string, Object> p_dicEventSetting)
        {
            SysEntity.TransResult m_TransResult = new SysEntity.TransResult();

            try
            {
                m_TransResult = g_DC.InsertIntoByEntity<SysEntity.EventSetting>(p_EmployeeEntity, "EventSetting", p_dicEventSetting);
            }
            catch (Exception ex)
            {
                m_TransResult.LogMessage = ex.Message;
                m_TransResult.isSuccess = false;
            }

            return m_TransResult;
        }

        public SysEntity.TransResult DelEventSetting(SysEntity.Employee p_EmployeeEntity, Dictionary<string, Object> p_dicEventSetting)
        {
            SysEntity.TransResult m_TransResult = new SysEntity.TransResult();
            Dictionary<string, Object> m_ArrayList = new Dictionary<string, Object>();

            SysEntity.EventSetting m_EventSetting = g_FH.ConvertEntity<SysEntity.EventSetting>(p_dicEventSetting, ref m_ArrayList);

            try
            {
                string m_SQL = "Delete [dbo].[EventID] Where [EventID]=@EventID";
                using (SqlCommand m_cmd = new SqlCommand(m_SQL))
                {
                    m_cmd.Parameters.Add(new SqlParameter("EventID", m_EventSetting.EventID));
                    m_TransResult = g_DC.ExecuteNonQuery(p_EmployeeEntity, m_cmd);
                }
            }
            catch (Exception ex)
            {
                m_TransResult.LogMessage = ex.Message;
                m_TransResult.isSuccess = false;
            }

            return m_TransResult;

        }

        public SysEntity.TransResult UpdEventSetting(SysEntity.Employee p_EmployeeEntity, Dictionary<string, Object> p_dicEventSetting)
        {
            SysEntity.TransResult m_TransResult = new SysEntity.TransResult();
            Dictionary<string, Object> m_ArrayList = new Dictionary<string, Object>();

            SysEntity.EventSetting m_EventSetting = g_FH.ConvertEntity<SysEntity.EventSetting>(p_dicEventSetting, ref m_ArrayList);

            try
            {
                string m_SQL = "";
                using (SqlCommand m_cmd = new SqlCommand())
                {
                    m_SQL = "Update [dbo].[EventSetting]  set EventModel =@EventModel , EventAction =@EventAction, EventDescription =@EventDescription, EventActionType =@EventActionType, EventFollow =@EventFollow, EventRef =@EventRef,ModifyUser=@ModifyUser , ModifyDate=(Convert(varchar(10),Getdate(),111)), ModifyTime=(Convert(varchar(8),Getdate(),108)) ";
                    m_cmd.Parameters.Add(new SqlParameter("EventModel", m_EventSetting.EventModel));
                    m_cmd.Parameters.Add(new SqlParameter("EventAction", m_EventSetting.EventAction));
                    m_cmd.Parameters.Add(new SqlParameter("EventDescription", m_EventSetting.EventDescription));
                    m_cmd.Parameters.Add(new SqlParameter("EventActionType", m_EventSetting.EventActionType));
                    m_cmd.Parameters.Add(new SqlParameter("EventFollow", m_EventSetting.EventFollow));
                    m_cmd.Parameters.Add(new SqlParameter("EventRef", m_EventSetting.EventRef));

                    m_cmd.Parameters.Add(new SqlParameter("ModifyUser", p_EmployeeEntity.WorkID));

                    if (m_EventSetting.EventID != "")
                    {
                        m_SQL += " where [EventID]=@EventID ";
                        m_cmd.Parameters.Add(new SqlParameter("EventID", m_EventSetting.EventID));
                        m_cmd.CommandText = m_SQL;
                        m_TransResult = g_DC.ExecuteNonQuery(p_EmployeeEntity, m_cmd);
                    }
                    else
                    {
                        m_TransResult.LogMessage = "EventID 不可為空!!!";
                        m_TransResult.isSuccess = false;
                    }

                }
            }
            catch (Exception ex)
            {
                m_TransResult.LogMessage = ex.Message;
                m_TransResult.isSuccess = false;
            }

            return m_TransResult;

        }


        #endregion

        #region Parameter Methods

        public SysEntity.TransResult GetParameters(SysEntity.Employee p_EmployeeEntity, Dictionary<string, Object> p_dicParameter)
        {
            SysEntity.TransResult m_TransResult = new SysEntity.TransResult();
            Dictionary<string, Object> m_objBetween = new Dictionary<string, Object>();
            SysEntity.Parameter m_Parameter = g_FH.ConvertEntity<SysEntity.Parameter>(p_dicParameter, ref m_objBetween);

            try
            {
                string m_SQL = "SELECT * FROM [dbo].[Parameter]   ";
                using (SqlCommand m_cmd = g_DC.GetQueryCommand(p_EmployeeEntity, m_SQL, m_Parameter, m_objBetween))
                {
                    m_TransResult = g_DC.GetDataTable(p_EmployeeEntity, m_cmd);
                }
            }
            catch (Exception ex)
            {
                m_TransResult.LogMessage = ex.Message;
                m_TransResult.isSuccess = false;
            }
            return m_TransResult;
        }

        public SysEntity.TransResult GetParameterGrid(SysEntity.Employee p_EmployeeEntity, Dictionary<string, Object> p_dicParameter, string p_PageIndex, string p_PageSize, Dictionary<string, Object> p_Param)
        {
            SysEntity.TransResult m_TransResult = new SysEntity.TransResult();
            Dictionary<string, Object> m_objBetween = new Dictionary<string, Object>();
            SysEntity.Parameter m_Parameter = g_FH.ConvertEntity<SysEntity.Parameter>(p_dicParameter, ref m_objBetween);
            ArrayList m_alGrid = new ArrayList();


            try
            {
                string m_SQL = "SELECT [ParamID],[Text],[Value],[Description] FROM [dbo].[Parameter]   ";
                using (SqlCommand m_cmd = g_DC.GetQueryCommand(p_EmployeeEntity, m_SQL, m_Parameter, m_objBetween))
                {

                    m_TransResult = g_DC.GetDataTableByGrid(p_EmployeeEntity, m_cmd, p_Param, p_PageSize, p_PageIndex);
                    if (m_TransResult.isSuccess)
                    {
                        m_TransResult = g_DC.GridDataCtrl(m_TransResult, p_PageSize, p_Param);
                    }
                }
            }
            catch (Exception ex)
            {
                m_TransResult.LogMessage = ex.Message;
                m_TransResult.isSuccess = false;
            }
            return m_TransResult;
        }

        public SysEntity.TransResult AddParameter(SysEntity.Employee p_EmployeeEntity, Dictionary<string, Object> p_dicParameter)
        {
            SysEntity.TransResult m_TransResult = new SysEntity.TransResult();

            try
            {
                m_TransResult = g_DC.InsertIntoByEntity<SysEntity.Parameter>(p_EmployeeEntity, "Parameter", p_dicParameter);
            }
            catch (Exception ex)
            {
                m_TransResult.LogMessage = ex.Message;
                m_TransResult.isSuccess = false;
            }

            return m_TransResult;

        }

        public SysEntity.TransResult DelParameter(SysEntity.Employee p_EmployeeEntity, Dictionary<string, Object> p_dicParameter)
        {
            SysEntity.TransResult m_TransResult = new SysEntity.TransResult();
            Dictionary<string, Object> m_ArrayList = new Dictionary<string, Object>();

            SysEntity.Parameter m_Parameter = g_FH.ConvertEntity<SysEntity.Parameter>(p_dicParameter, ref m_ArrayList);

            try
            {
                string m_SQL = "Delete [dbo].[Parameter] Where [ParamID]=@ParamID";
                using (SqlCommand m_cmd = new SqlCommand(m_SQL))
                {
                    m_cmd.Parameters.Add(new SqlParameter("ParamID", m_Parameter.ParamID));
                    m_TransResult = g_DC.ExecuteNonQuery(p_EmployeeEntity, m_cmd);
                }
            }
            catch (Exception ex)
            {
                m_TransResult.LogMessage = ex.Message;
                m_TransResult.isSuccess = false;
            }

            return m_TransResult;

        }

        public SysEntity.TransResult UpdParameter(SysEntity.Employee p_EmployeeEntity, Dictionary<string, Object> p_dicParameter)
        {
            SysEntity.TransResult m_TransResult = new SysEntity.TransResult();
            Dictionary<string, Object> m_ArrayList = new Dictionary<string, Object>();

            SysEntity.Parameter m_Parameter = g_FH.ConvertEntity<SysEntity.Parameter>(p_dicParameter, ref m_ArrayList);

            try
            {
                string m_SQL = "";
                using (SqlCommand m_cmd = new SqlCommand())
                {
                    m_SQL = "Update [dbo].[Parameter]  set Text =@Text , Value =@Value, Description =@Description,ModifyUser=@ModifyUser , ModifyDate=(Convert(varchar(10),Getdate(),111)), ModifyTime=(Convert(varchar(8),Getdate(),108)) ";
                    m_cmd.Parameters.Add(new SqlParameter("Text", m_Parameter.Text));
                    m_cmd.Parameters.Add(new SqlParameter("Value", m_Parameter.Value));
                    m_cmd.Parameters.Add(new SqlParameter("ParentParamID", m_Parameter.ParentParamID));
                    m_cmd.Parameters.Add(new SqlParameter("Description", m_Parameter.Description));
                    m_cmd.Parameters.Add(new SqlParameter("ModifyUser", p_EmployeeEntity.WorkID));

                    if (m_Parameter.ParamID != "")
                    {
                        m_SQL += " where [ParamID]=@ParamID ";
                        m_cmd.Parameters.Add(new SqlParameter("ParamID", m_Parameter.ParamID));
                        m_cmd.CommandText = m_SQL;
                        m_TransResult = g_DC.ExecuteNonQuery(p_EmployeeEntity, m_cmd);
                    }
                    else
                    {
                        m_TransResult.LogMessage = "ParamID 不可為空!!!";
                        m_TransResult.isSuccess = false;
                    }

                }
            }
            catch (Exception ex)
            {
                m_TransResult.LogMessage = ex.Message;
                m_TransResult.isSuccess = false;
            }

            return m_TransResult;

        }

        public SysEntity.TransResult GetParameter(SysEntity.Employee p_EmployeeEntity, SysEntity.Parameter p_Parameter)
        {
            SysEntity.TransResult m_TransResult = new SysEntity.TransResult();
            try
            {
                string m_SQL = "SELECT [ParamID],[Text],[Value],[Description] FROM [dbo].[Parameter]   ";
                using (SqlCommand m_cmd = g_DC.GetQueryCommand(p_EmployeeEntity, m_SQL, p_Parameter))
                {
                    m_TransResult = g_DC.GetDataTable(p_EmployeeEntity, m_cmd);
                }
            }
            catch (Exception ex)
            {
                m_TransResult.LogMessage = ex.Message;
                m_TransResult.isSuccess = false;
            }
            return m_TransResult;
        }


        public SysEntity.TransResult GetSelSectionType(SysEntity.Employee p_EmployeeEntity, Dictionary<string, Object> p_dicParameter)
        {
            SysEntity.TransResult m_TransResult = new SysEntity.TransResult();
            Dictionary<string, Object> m_objBetween = new Dictionary<string, Object>();
            SysEntity.Parameter m_Parameter = g_FH.ConvertEntity<SysEntity.Parameter>(p_dicParameter, ref m_objBetween);

            try
            {
                string m_SQL = "SELECT [ParamID],[Text] SectionType,[Description] FROM [dbo].[Parameter] where [ParamID] ='SectionType' ";
               
                using (SqlCommand m_cmd = g_DC.GetQueryCommand(p_EmployeeEntity, m_SQL, m_Parameter, m_objBetween))
                {
                    m_TransResult = g_DC.GetDataTable(p_EmployeeEntity, m_cmd);
                }
            }
            catch (Exception ex)
            {
                m_TransResult.LogMessage = ex.Message;
                m_TransResult.isSuccess = false;
            }
            return m_TransResult;
        }

        public SysEntity.TransResult GetSectionType(SysEntity.Employee p_EmployeeEntity, Dictionary<string, Object> p_dicParameter, string p_PageIndex, string p_PageSize, Dictionary<string, Object> p_Param)
        {
            SysEntity.TransResult m_TransResult = new SysEntity.TransResult();
            Dictionary<string, Object> m_objBetween = new Dictionary<string, Object>();
            SysEntity.Parameter m_Parameter = g_FH.ConvertEntity<SysEntity.Parameter>(p_dicParameter, ref m_objBetween);
            ArrayList m_alGrid = new ArrayList();

            string[] m_LikeProp = { "Description" };
            try
            {
                string m_SQL = "SELECT [ParamID],[Text] SectionType,[Description] FROM [dbo].[Parameter] where [ParamID] ='SectionType' ";
                using (SqlCommand m_cmd = g_DC.GetQueryCommand(p_EmployeeEntity, m_SQL, m_Parameter, m_objBetween, m_LikeProp))
                {

                    m_TransResult = g_DC.GetDataTableByGrid(p_EmployeeEntity, m_cmd, p_Param, p_PageSize, p_PageIndex);
                    if (m_TransResult.isSuccess)
                    {
                        m_TransResult = g_DC.GridDataCtrl(m_TransResult, p_PageSize, p_Param);
                    }
                }
            }
            catch (Exception ex)
            {
                m_TransResult.LogMessage = ex.Message;
                m_TransResult.isSuccess = false;
            }
            return m_TransResult;
        }

        public SysEntity.TransResult GetSectionID(SysEntity.Employee p_EmployeeEntity, Dictionary<string, Object> p_dicParameter, string p_PageIndex, string p_PageSize, Dictionary<string, Object> p_Param)
        {
            SysEntity.TransResult m_TransResult = new SysEntity.TransResult();
            Dictionary<string, Object> m_objBetween = new Dictionary<string, Object>();
            SysEntity.Parameter m_Parameter = g_FH.ConvertEntity<SysEntity.Parameter>(p_dicParameter, ref m_objBetween);
            ArrayList m_alGrid = new ArrayList();

            string[] m_LikeProp = { "Description" };
            try
            {
                string m_SQL = "SELECT [ParamID],[Text] SectionID,[Description] FROM [dbo].[Parameter] where [ParamID] ='SectionID' ";
                using (SqlCommand m_cmd = g_DC.GetQueryCommand(p_EmployeeEntity, m_SQL, m_Parameter, m_objBetween, m_LikeProp))
                {

                    m_TransResult = g_DC.GetDataTableByGrid(p_EmployeeEntity, m_cmd, p_Param, p_PageSize, p_PageIndex);
                    if (m_TransResult.isSuccess)
                    {
                        m_TransResult = g_DC.GridDataCtrl(m_TransResult, p_PageSize, p_Param);
                    }
                }
            }
            catch (Exception ex)
            {
                m_TransResult.LogMessage = ex.Message;
                m_TransResult.isSuccess = false;
            }
            return m_TransResult;
        }

        public DataTable GetParameterByParamID(SysEntity.Employee p_EmployeeEntity, Dictionary<string, Object> p_dicParam)
        {
            DataTable m_dt = new DataTable();

            Dictionary<string, Object> m_ArrayList = new Dictionary<string, Object>();

            SysEntity.Parameter m_Parameter = g_FH.ConvertEntity<SysEntity.Parameter>(p_dicParam, ref m_ArrayList);


            string m_SQL = "Select Text " + m_Parameter.ParamID + ",Description  from [dbo].[Parameter]  ";
            using (SqlCommand m_cmd = new SqlCommand(m_SQL))
            {


                SysEntity.TransResult m_TransResult = new SysEntity.TransResult();
                m_cmd.Parameters.Add(new SqlParameter("ParamID", m_Parameter.ParamID));
                m_TransResult = g_DC.GetDataTable(p_EmployeeEntity, m_cmd);
                if (m_TransResult.isSuccess)
                {
                    m_dt = (DataTable)m_TransResult.ResultEntity;
                }
            }
            return m_dt;
        }

        public SysEntity.TransResult GetControlType(SysEntity.Employee p_EmployeeEntity, Dictionary<string, Object> p_dicParameter, string p_PageIndex, string p_PageSize, Dictionary<string, Object> p_Param)
        {
            SysEntity.TransResult m_TransResult = new SysEntity.TransResult();
            Dictionary<string, Object> m_objBetween = new Dictionary<string, Object>();
            SysEntity.Parameter m_Parameter = g_FH.ConvertEntity<SysEntity.Parameter>(p_dicParameter, ref m_objBetween);
            ArrayList m_alGrid = new ArrayList();
            string[] m_LikeProp = { "Description" };


            try
            {
                string m_SQL = "SELECT [ParamID],[Text] ControlType,[Description] FROM [dbo].[Parameter] where [ParamID] ='ControlType' ";
                using (SqlCommand m_cmd = g_DC.GetQueryCommand(p_EmployeeEntity, m_SQL, m_Parameter, m_objBetween, m_LikeProp))
                {

                    m_TransResult = g_DC.GetDataTableByGrid(p_EmployeeEntity, m_cmd, p_Param, p_PageSize, p_PageIndex);
                    if (m_TransResult.isSuccess)
                    {
                        m_TransResult = g_DC.GridDataCtrl(m_TransResult, p_PageSize, p_Param);
                    }
                }
            }
            catch (Exception ex)
            {
                m_TransResult.LogMessage = ex.Message;
                m_TransResult.isSuccess = false;
            }
            return m_TransResult;
        }

        #endregion

        #region QueryForm Methods


        public SysEntity.TransResult GetQueryForms(SysEntity.Employee p_EmployeeEntity, Dictionary<string, Object> p_dicQueryForm)
        {

            SysEntity.TransResult m_TransResult = new SysEntity.TransResult();
            Dictionary<string, Object> m_objBetween = new Dictionary<string, Object>();
            SysEntity.QueryForm m_QueryForm = g_FH.ConvertEntity<SysEntity.QueryForm>(p_dicQueryForm, ref m_objBetween);

            try
            {
                string m_SQL = "SELECT ''ViewDetailID,''ValidParam, 'General1' SectionType,'' ValidID,'' EventAction,'GridData' EventActionType,'' ValidSource,'0'ModifyKey ,''EventFollow,QF.* FROM [dbo].[QueryForm] QF   ";
                using (SqlCommand m_cmd = g_DC.GetQueryCommand(p_EmployeeEntity, m_SQL, p_dicQueryForm, m_objBetween))
                {
                    m_TransResult = g_DC.GetDataTable(p_EmployeeEntity, m_cmd);
                }
            }
            catch (Exception ex)
            {
                m_TransResult.LogMessage = ex.Message;
                m_TransResult.isSuccess = false;
            }
            return m_TransResult;
        }

        public SysEntity.TransResult GetQueryFormGrid(SysEntity.Employee p_EmployeeEntity, Dictionary<string, Object> p_dicQueryForm, string p_PageIndex, string p_PageSize, Dictionary<string, Object> p_Param)
        {
            SysEntity.TransResult m_TransResult = new SysEntity.TransResult();
            Dictionary<string, Object> m_objBetween = new Dictionary<string, Object>();
            SysEntity.QueryForm m_QueryForm = g_FH.ConvertEntity<SysEntity.QueryForm>(p_dicQueryForm, ref m_objBetween);
            ArrayList m_alGrid = new ArrayList();


            try
            {
                string m_SQL = "SELECT [QueryID],[SectionType],[ControlType],[FieldID] FROM [dbo].[QueryForm]   ";
                using (SqlCommand m_cmd = g_DC.GetQueryCommand(p_EmployeeEntity, m_SQL, m_QueryForm, m_objBetween))
                {

                    m_TransResult = g_DC.GetDataTableByGrid(p_EmployeeEntity, m_cmd, p_Param, p_PageSize, p_PageIndex);
                    if (m_TransResult.isSuccess)
                    {
                        m_TransResult = g_DC.GridDataCtrl(m_TransResult, p_PageSize, p_Param);
                    }
                }
            }
            catch (Exception ex)
            {
                m_TransResult.LogMessage = ex.Message;
                m_TransResult.isSuccess = false;
            }
            return m_TransResult;
        }

        public SysEntity.TransResult AddQueryForm(SysEntity.Employee p_EmployeeEntity, Dictionary<string, Object> p_dicQueryForm)
        {
            SysEntity.TransResult m_TransResult = new SysEntity.TransResult();
            try
            {
                m_TransResult = g_DC.InsertIntoByEntity<SysEntity.QueryForm>(p_EmployeeEntity, "QueryForm", p_dicQueryForm);
            }
            catch (Exception ex)
            {
                m_TransResult.LogMessage = ex.Message;
                m_TransResult.isSuccess = false;
            }

            return m_TransResult;

        }

        public SysEntity.TransResult DelQueryForm(SysEntity.Employee p_EmployeeEntity, Dictionary<string, Object> p_dicQueryForm)
        {
            SysEntity.TransResult m_TransResult = new SysEntity.TransResult();
            Dictionary<string, Object> m_ArrayList = new Dictionary<string, Object>();

            SysEntity.QueryForm m_QueryForm = g_FH.ConvertEntity<SysEntity.QueryForm>(p_dicQueryForm, ref m_ArrayList);

            try
            {
                string m_SQL = "Delete [dbo].[QueryForm] Where [QueryID]=@QueryID";
                using (SqlCommand m_cmd = new SqlCommand(m_SQL))
                {
                    m_cmd.Parameters.Add(new SqlParameter("QueryID", m_QueryForm.QueryID));
                    m_TransResult = g_DC.ExecuteNonQuery(p_EmployeeEntity, m_cmd);
                }
            }
            catch (Exception ex)
            {
                m_TransResult.LogMessage = ex.Message;
                m_TransResult.isSuccess = false;
            }

            return m_TransResult;

        }

        public SysEntity.TransResult UpdQueryForm(SysEntity.Employee p_EmployeeEntity, Dictionary<string, Object> p_dicQueryForm)
        {
            SysEntity.TransResult m_TransResult = new SysEntity.TransResult();
            Dictionary<string, Object> m_ArrayList = new Dictionary<string, Object>();

            SysEntity.QueryForm m_QueryForm = g_FH.ConvertEntity<SysEntity.QueryForm>(p_dicQueryForm, ref m_ArrayList);

            try
            {
                string m_SQL = "";
                using (SqlCommand m_cmd = new SqlCommand())
                {
                    m_SQL = "Update [dbo].[QueryForm]  set SectionType =@SectionType ,ControlType =@ControlType,FieldID =@FieldID,ControlRef =@ControlRef,Height =@Height,Width =@Width,EventID =@EventID,MaxLength =@MaxLength,ModifyUser=@ModifyUser , ModifyDate=(Convert(varchar(10),Getdate(),111)), ModifyTime=(Convert(varchar(8),Getdate(),108)) ";
                    m_cmd.Parameters.Add(new SqlParameter("SectionType", m_QueryForm.SectionType));
                    m_cmd.Parameters.Add(new SqlParameter("ControlType", m_QueryForm.ControlType));
                    m_cmd.Parameters.Add(new SqlParameter("FieldID", m_QueryForm.FieldID));
                    m_cmd.Parameters.Add(new SqlParameter("ControlRef", m_QueryForm.ControlRef));
                    m_cmd.Parameters.Add(new SqlParameter("Height", m_QueryForm.Height));
                    m_cmd.Parameters.Add(new SqlParameter("Width", m_QueryForm.Width));
                    m_cmd.Parameters.Add(new SqlParameter("EventID", m_QueryForm.EventID));
                    m_cmd.Parameters.Add(new SqlParameter("MaxLength", m_QueryForm.MaxLength));

                    m_cmd.Parameters.Add(new SqlParameter("ModifyUser", p_EmployeeEntity.WorkID));

                    if (m_QueryForm.QueryID != "")
                    {
                        m_SQL += " where [QueryID]=@QueryID ";
                        m_cmd.Parameters.Add(new SqlParameter("QueryID", m_QueryForm.QueryID));
                        m_cmd.CommandText = m_SQL;
                        m_TransResult = g_DC.ExecuteNonQuery(p_EmployeeEntity, m_cmd);
                    }
                    else
                    {
                        m_TransResult.LogMessage = "QueryID 不可為空!!!";
                        m_TransResult.isSuccess = false;
                    }

                }
            }
            catch (Exception ex)
            {
                m_TransResult.LogMessage = ex.Message;
                m_TransResult.isSuccess = false;
            }

            return m_TransResult;

        }

        public SysEntity.TransResult GetQueryFormGroupByQueryID(SysEntity.Employee p_EmployeeEntity, Dictionary<string, Object> p_dicQueryForm, string p_PageIndex, string p_PageSize, Dictionary<string, Object> p_Param)
        {
            SysEntity.TransResult m_TransResult = new SysEntity.TransResult();
            Dictionary<string, Object> m_objBetween = new Dictionary<string, Object>();
            SysEntity.QueryForm m_QueryForm = g_FH.ConvertEntity<SysEntity.QueryForm>(p_dicQueryForm, ref m_objBetween);
            ArrayList m_alGrid = new ArrayList();


            try
            {
                string m_SQL = "SELECT distinct [QueryID],[FormTitle],[Description] FROM [dbo].[QueryForm]   ";
                using (SqlCommand m_cmd = g_DC.GetQueryCommand(p_EmployeeEntity, m_SQL, m_QueryForm, m_objBetween))
                {

                    m_TransResult = g_DC.GetDataTableByGrid(p_EmployeeEntity, m_cmd, p_Param, p_PageSize, p_PageIndex);
                    if (m_TransResult.isSuccess)
                    {
                        m_TransResult = g_DC.GridDataCtrl(m_TransResult, p_PageSize, p_Param);
                    }
                }
            }
            catch (Exception ex)
            {
                m_TransResult.LogMessage = ex.Message;
                m_TransResult.isSuccess = false;
            }
            return m_TransResult;
        }


        public string GetEventIDByQueryID(SysEntity.Employee p_EmployeeEntity, string p_QueryID)
        {
            string m_EventID = "";
            DataTable m_dt = new DataTable();
            string m_SQL = "SELECT [EventID] FROM [dbo].[QueryForm] where [FieldID]='Search' and QueryID=@QueryID  ";
            using (SqlCommand m_cmd = new SqlCommand(m_SQL))
            {
                m_cmd.Parameters.Add(new SqlParameter("QueryID", p_QueryID));

                SysEntity.TransResult m_TransResult = new SysEntity.TransResult();


                m_TransResult = g_DC.GetDataTable(p_EmployeeEntity, m_cmd);
                if (m_TransResult.isSuccess)
                {
                    m_dt = (DataTable)m_TransResult.ResultEntity;
                    foreach (DataRow dr in m_dt.Rows)
                    {
                        m_EventID = dr["EventID"].ToString();
                    }
                }
            }
            return m_EventID;
        }


        #endregion

        #region TransferJobSetting Methods

        public SysEntity.TransResult GetTransferJobSettings(SysEntity.Employee p_EmployeeEntity, Dictionary<string, Object> p_dicTransferJobSetting)
        {
            SysEntity.TransResult m_TransResult = new SysEntity.TransResult();
            Dictionary<string, Object> m_objBetween = new Dictionary<string, Object>();
            SysEntity.TransferJobSetting m_TransferJobSetting = g_FH.ConvertEntity<SysEntity.TransferJobSetting>(p_dicTransferJobSetting, ref m_objBetween);

            try
            {
                //string m_SQL = "SELECT * FROM [dbo].[TransferJobSetting]   ";
                string m_SQL = "SELECT  [TransferID] ,'Transfer'+ [TransferID]+'-Desc' as [TransferDesc],[TransferEventID],[TransferCycle],[TransferUnit],[TransferParam],[TransferOwner],[TransferStatus],[CreateDate],[CreateTime],[CreateUser],[ModifyUser],[ModifyDate],[ModifyTime] FROM [dbo].[TransferJobSetting]   ";
                using (SqlCommand m_cmd = g_DC.GetQueryCommand(p_EmployeeEntity, m_SQL, m_TransferJobSetting, m_objBetween))
                {
                    m_TransResult = g_DC.GetDataTable(p_EmployeeEntity, m_cmd);
                }
            }
            catch (Exception ex)
            {
                m_TransResult.LogMessage = ex.Message;
                m_TransResult.isSuccess = false;
            }
            return m_TransResult;
        }

        public SysEntity.TransResult GetTransferJobSettingGrid(SysEntity.Employee p_EmployeeEntity, Dictionary<string, Object> p_dicTransferJobSetting, string p_PageIndex, string p_PageSize, Dictionary<string, Object> p_Param)
        {
            SysEntity.TransResult m_TransResult = new SysEntity.TransResult();
            Dictionary<string, Object> m_objBetween = new Dictionary<string, Object>();
            SysEntity.TransferJobSetting m_TransferJobSetting = g_FH.ConvertEntity<SysEntity.TransferJobSetting>(p_dicTransferJobSetting, ref m_objBetween);
            ArrayList m_alGrid = new ArrayList();


            try
            {
                string[] m_LikeProp = { "TransferDesc" };

                //string m_SQL = "SELECT [TransferID],[TransferDesc],[TransferCycle],[TransferUnit],[TransferStatus] FROM [dbo].[TransferJobSetting]   ";
                string m_SQL = "SELECT  [TransferID] ,'Transfer'+ [TransferID]+'-Desc' as [TransferDesc],[TransferCycle],[TransferUnit],[TransferStatus] FROM [dbo].[TransferJobSetting]   ";


                using (SqlCommand m_cmd = g_DC.GetQueryCommand(p_EmployeeEntity, m_SQL, m_TransferJobSetting, m_objBetween, m_LikeProp))
                {

                    m_TransResult = g_DC.GetDataTableByGrid(p_EmployeeEntity, m_cmd, p_Param, p_PageSize, p_PageIndex);
                    if (m_TransResult.isSuccess)
                    {
                        m_TransResult = g_DC.GridDataCtrl(m_TransResult, p_PageSize, p_Param);
                    }
                }
            }
            catch (Exception ex)
            {
                m_TransResult.LogMessage = ex.Message;
                m_TransResult.isSuccess = false;
            }
            return m_TransResult;
        }

        public SysEntity.TransResult AddTransferJobSetting(SysEntity.Employee p_EmployeeEntity, Dictionary<string, Object> p_dicTransferJobSetting)
        {
            SysEntity.TransResult m_TransResult = new SysEntity.TransResult();
            try
            {
                m_TransResult = g_DC.InsertIntoByEntity<SysEntity.TransferJobSetting>(p_EmployeeEntity, "TransferJobSetting", p_dicTransferJobSetting);
            }
            catch (Exception ex)
            {
                m_TransResult.LogMessage = ex.Message;
                m_TransResult.isSuccess = false;
            }

            return m_TransResult;

        }

        public SysEntity.TransResult DelTransferJobSetting(SysEntity.Employee p_EmployeeEntity, Dictionary<string, Object> p_dicTransferJobSetting)
        {
            SysEntity.TransResult m_TransResult = new SysEntity.TransResult();
            Dictionary<string, Object> m_ArrayList = new Dictionary<string, Object>();

            SysEntity.TransferJobSetting m_TransferJobSetting = g_FH.ConvertEntity<SysEntity.TransferJobSetting>(p_dicTransferJobSetting, ref m_ArrayList);

            try
            {
                string m_SQL = "Delete [dbo].[TransferID] Where [TransferJobSettingID]=@TransferID";
                using (SqlCommand m_cmd = new SqlCommand(m_SQL))
                {
                    m_cmd.Parameters.Add(new SqlParameter("TransferJobSettingID", m_TransferJobSetting.TransferID));
                    m_TransResult = g_DC.ExecuteNonQuery(p_EmployeeEntity, m_cmd);
                }
            }
            catch (Exception ex)
            {
                m_TransResult.LogMessage = ex.Message;
                m_TransResult.isSuccess = false;
            }

            return m_TransResult;

        }

        public SysEntity.TransResult UpdTransferJobSetting(SysEntity.Employee p_EmployeeEntity, Dictionary<string, Object> p_dicTransferJobSetting)
        {
            SysEntity.TransResult m_TransResult = new SysEntity.TransResult();
            Dictionary<string, Object> m_ArrayList = new Dictionary<string, Object>();

            SysEntity.TransferJobSetting m_TransferJobSetting = g_FH.ConvertEntity<SysEntity.TransferJobSetting>(p_dicTransferJobSetting, ref m_ArrayList);

            try
            {
                string m_SQL = "";
                using (SqlCommand m_cmd = new SqlCommand())
                {
                    m_SQL = "Update [dbo].[TransferJobSetting]  set TransferDesc =@TransferDesc ,TransferCycle =@TransferCycle ,TransferUnit =@TransferUnit ,TransferParam =@TransferParam ,TransferOwner =@TransferOwner ,TransferStatus =@TransferStatus ,ModifyUser=@ModifyUser, ModifyDate=(Convert(varchar(10),Getdate(),111)), ModifyTime=(Convert(varchar(8),Getdate(),108))  ";
                    m_cmd.Parameters.Add(new SqlParameter("TransferDesc", m_TransferJobSetting.TransferDesc));
                    m_cmd.Parameters.Add(new SqlParameter("TransferEventID", m_TransferJobSetting.TransferEventID));
                    m_cmd.Parameters.Add(new SqlParameter("TransferCycle", m_TransferJobSetting.TransferCycle));
                    m_cmd.Parameters.Add(new SqlParameter("TransferUnit", m_TransferJobSetting.TransferUnit));
                    m_cmd.Parameters.Add(new SqlParameter("TransferParam", m_TransferJobSetting.TransferParam));
                    m_cmd.Parameters.Add(new SqlParameter("TransferOwner", m_TransferJobSetting.TransferOwner));
                    m_cmd.Parameters.Add(new SqlParameter("TransferStatus", m_TransferJobSetting.TransferStatus));
                    m_cmd.Parameters.Add(new SqlParameter("ModifyUser", p_EmployeeEntity.WorkID));

                    if (m_TransferJobSetting.TransferID != "")
                    {
                        m_SQL += " where [TransferID]=@TransferID ";
                        m_cmd.Parameters.Add(new SqlParameter("TransferID", m_TransferJobSetting.TransferID));
                        m_cmd.CommandText = m_SQL;
                        m_TransResult = g_DC.ExecuteNonQuery(p_EmployeeEntity, m_cmd);
                    }
                    else
                    {
                        m_TransResult.LogMessage = "TransferID 不可為空!!!";
                        m_TransResult.isSuccess = false;
                    }

                }
            }
            catch (Exception ex)
            {
                m_TransResult.LogMessage = ex.Message;
                m_TransResult.isSuccess = false;
            }

            return m_TransResult;

        }

        public SysEntity.TransResult AddTransferJobLog(SysEntity.Employee p_EmployeeEntity, SysEntity.TransferJobLog p_TransferJobLog)
        {
            SysEntity.TransResult m_TransResult = new SysEntity.TransResult();
            try
            {
                Dictionary<string, Object> m_TransferJobLog =
                p_TransferJobLog.GetType()
     .GetProperties(BindingFlags.Instance | BindingFlags.Public)
          .ToDictionary(prop => prop.Name, prop => prop.GetValue(p_TransferJobLog, null));
                m_TransResult = g_DC.InsertIntoByEntity<SysEntity.TransferJobLog>(p_EmployeeEntity, "TransferJobLog", m_TransferJobLog, false);
            }
            catch (Exception ex)
            {
                m_TransResult.LogMessage = ex.Message;
                m_TransResult.isSuccess = false;
            }

            return m_TransResult;

        }

        public SysEntity.TransResult GetTransferJobLog(SysEntity.Employee p_EmployeeEntity, Dictionary<string, Object> p_dicTransferJobLog)
        {
            SysEntity.TransResult m_TransResult = new SysEntity.TransResult();
            Dictionary<string, Object> m_objBetween = new Dictionary<string, Object>();
            SysEntity.TransferJobLog m_TransferJobLog = g_FH.ConvertEntity<SysEntity.TransferJobLog>(p_dicTransferJobLog, ref m_objBetween);

            try
            {
                string m_SQL = "SELECT * FROM [dbo].[TransferJobLog] where TransferDate between CONVERT (CHAR , DATEADD(day,-30,GETDATE())  , 111)  and CONVERT (CHAR , GETDATE( )  , 111)   ";
                using (SqlCommand m_cmd = g_DC.GetQueryCommand(p_EmployeeEntity, m_SQL, m_TransferJobLog, m_objBetween, null, " order by [TransferDate] Desc,[TransferTime] Desc"))
                {
                    m_TransResult = g_DC.GetDataTable(p_EmployeeEntity, m_cmd);
                }
            }
            catch (Exception ex)
            {
                m_TransResult.LogMessage = ex.Message;
                m_TransResult.isSuccess = false;
            }
            return m_TransResult;
        }



        #endregion

       


        #region OwnerMaintain Methods


        public SysEntity.TransResult GetOwnerMaintain(SysEntity.Employee p_EmployeeEntity, Dictionary<string, Object> p_dicOwnerMaintain)
        {
            SysEntity.TransResult m_TransResult = new SysEntity.TransResult();
            Dictionary<string, Object> m_objBetween = new Dictionary<string, Object>();
           
            try
            {
                string m_SQL = "SELECT ParamValue1 Site,ParamValue2 System1,ParamValue3 System3,ParamValue2,ParamValue3 ,CaseOwner,Deputy1,Deputy2,dbo.fun_GetEmployeeName(CaseOwner)EmployeeName,AgentDateS,AgentTimeS,AgentDateE,AgentTimeE FROM [dbo].[OwnerMaintain]   ";

                m_SQL += " where ParamValue1 in (select Description from [dbo].[Site] where SiteID='" + p_EmployeeEntity.CompanyCode+ "') ";

                using (SqlCommand m_cmd = g_DC.GetQueryCommand(p_EmployeeEntity, m_SQL, p_dicOwnerMaintain, m_objBetween))
                {
                    m_TransResult = g_DC.GetDataTable(p_EmployeeEntity, m_cmd);
                }
            }
            catch (Exception ex)
            {
                m_TransResult.LogMessage = ex.Message;
                m_TransResult.isSuccess = false;
            }
            return m_TransResult;
        }

        public SysEntity.TransResult GetOwnerMaintainGrid(SysEntity.Employee p_EmployeeEntity, Dictionary<string, Object> p_dicOwnerMaintain, string p_PageIndex, string p_PageSize, Dictionary<string, Object> p_Param)
        {
            SysEntity.TransResult m_TransResult = new SysEntity.TransResult();
            Dictionary<string, Object> m_objBetween = new Dictionary<string, Object>();
            SysEntity.OwnerMaintain m_OwnerMaintain = g_FH.ConvertEntity<SysEntity.OwnerMaintain>(p_dicOwnerMaintain, ref m_objBetween);
            ArrayList m_alGrid = new ArrayList();
            string[] m_LikeProp = { "CaseOwner", "CustKeyWords", "Deputy1", "Deputy2" };


            try
            {
                string m_SQL = "SELECT ParamValue1 Site,MP1.sTextTW System1Text,MP.sTextTW System3Text,dbo.fun_GetEmployeeName(CaseOwner)CaseOwner,dbo.fun_GetEmployeeName(Deputy1)Deputy1,dbo.fun_GetEmployeeName(Deputy2)Deputy2,ParamValue2 System1,ParamValue3 System3,ParamValue2,ParamValue3  ";
                m_SQL += " ,dbo.fun_GetEmployeeName(O.[ModifyUser])ModifyUser ,O.[ModifyDate]  FROM [dbo].[OwnerMaintain] O ";
                m_SQL += " LEFT OUTER JOIN dbo.MultiParameter MP ON O.ParamValue3 = MP.sValue  ";
                       m_SQL += "  LEFT OUTER JOIN dbo.MultiParameter MP1 ON O.ParamValue2 = MP1.sValue  ";
                       using (SqlCommand m_cmd = g_DC.GetQueryCommand(p_EmployeeEntity, m_SQL, m_OwnerMaintain, m_objBetween, m_LikeProp))
                {

                    m_TransResult = g_DC.GetDataTableByGrid(p_EmployeeEntity, m_cmd, p_Param, p_PageSize, p_PageIndex);
                    if (m_TransResult.isSuccess)
                    {
                        m_TransResult = g_DC.GridDataCtrl(m_TransResult, p_PageSize, p_Param);
                    }
                }
            }
            catch (Exception ex)
            {
                m_TransResult.LogMessage = ex.Message;
                m_TransResult.isSuccess = false;
            }
            return m_TransResult;
        }

        public SysEntity.TransResult AddOwnerMaintain(SysEntity.Employee p_EmployeeEntity, Dictionary<string, Object> p_dicOwnerMaintain)
        {
            SysEntity.TransResult m_TransResult = new SysEntity.TransResult();

            try
            {
                string m_SQL = "";
                Dictionary<string, Object> m_ArrayList = new Dictionary<string, Object>();

                SysEntity.OwnerMaintain m_OwnerMaintain = g_FH.ConvertEntity<SysEntity.OwnerMaintain>(p_dicOwnerMaintain, ref m_ArrayList);

                using (SqlCommand m_cmd = new SqlCommand())
                {
                    m_SQL = "Insert into [dbo].[OwnerMaintain] ( [OwnerType],[ParamValue1],[ParamValue2],[ParamValue3],[CaseOwner],[Deputy1],[Deputy2],[AgentDateS],[AgentTimeS],[AgentDateE],[AgentTimeE]) ";
                    m_SQL += " Values(@OwnerType,@ParamValue1,@ParamValue2,@ParamValue3,@CaseOwner,@Deputy1,@Deputy2,@AgentDateS,@AgentTimeS,@AgentDateE,@AgentTimeE) ";

                    m_cmd.Parameters.Add(new SqlParameter("OwnerType", "ISSUE"));
                    m_cmd.Parameters.Add(new SqlParameter("ParamValue1", m_OwnerMaintain.Site));
                    m_cmd.Parameters.Add(new SqlParameter("ParamValue2", p_dicOwnerMaintain["System1"].ToString()));
                    m_cmd.Parameters.Add(new SqlParameter("ParamValue3", p_dicOwnerMaintain["System3"].ToString()));
                    m_cmd.Parameters.Add(new SqlParameter("CaseOwner", m_OwnerMaintain.CaseOwner));
                    m_cmd.Parameters.Add(new SqlParameter("Deputy1", m_OwnerMaintain.Deputy1));
                    m_cmd.Parameters.Add(new SqlParameter("Deputy2", m_OwnerMaintain.Deputy2));
                    m_cmd.Parameters.Add(new SqlParameter("AgentDateS", m_OwnerMaintain.AgentDateS));
                    m_cmd.Parameters.Add(new SqlParameter("AgentTimeS", m_OwnerMaintain.AgentTimeS));
                    m_cmd.Parameters.Add(new SqlParameter("AgentDateE", m_OwnerMaintain.AgentDateE));
                    m_cmd.Parameters.Add(new SqlParameter("AgentTimeE", m_OwnerMaintain.AgentTimeE));

                    m_cmd.CommandText = m_SQL;
                    m_TransResult = g_DC.ExecuteNonQuery(p_EmployeeEntity, m_cmd);


                }
            }
            catch (Exception ex)
            {
                m_TransResult.LogMessage = ex.Message;
                m_TransResult.isSuccess = false;
            }

            return m_TransResult;

        }

        public SysEntity.TransResult UpdOwnerMaintain(SysEntity.Employee p_EmployeeEntity, Dictionary<string, Object> p_dicOwnerMaintain)
        {
            SysEntity.TransResult m_TransResult = new SysEntity.TransResult();
            Dictionary<string, Object> m_ArrayList = new Dictionary<string, Object>();

            SysEntity.OwnerMaintain m_OwnerMaintain = g_FH.ConvertEntity<SysEntity.OwnerMaintain>(p_dicOwnerMaintain, ref m_ArrayList);

            try
            {

                string m_SQL = "";
                using (SqlCommand m_cmd = new SqlCommand())
                {
                    m_SQL = "Update [dbo].[OwnerMaintain]  set CaseOwner =@CaseOwner ,Deputy1 = @Deputy1,Deputy2 = @Deputy2,AgentDateS =@AgentDateS, AgentTimeS =@AgentTimeS,AgentDateE =@AgentDateE, AgentTimeE =@AgentTimeE,ModifyUser=@ModifyUser , ModifyDate=(Convert(varchar(10),Getdate(),111)), ModifyTime=(Convert(varchar(8),Getdate(),108)) ";
                    m_SQL += " where  OwnerType=@OwnerType and ParamValue1=@ParamValue1 and ParamValue2=@ParamValue2 and ParamValue3=@ParamValue3 ";
                   
                    m_cmd.Parameters.Add(new SqlParameter("CaseOwner", m_OwnerMaintain.CaseOwner));
                    m_cmd.Parameters.Add(new SqlParameter("Deputy1", m_OwnerMaintain.Deputy1));
                    m_cmd.Parameters.Add(new SqlParameter("Deputy2", m_OwnerMaintain.Deputy2));
                    m_cmd.Parameters.Add(new SqlParameter("AgentDateS", m_OwnerMaintain.AgentDateS));
                    m_cmd.Parameters.Add(new SqlParameter("AgentTimeS", m_OwnerMaintain.AgentTimeS));
                    m_cmd.Parameters.Add(new SqlParameter("AgentDateE", m_OwnerMaintain.AgentDateE));
                    m_cmd.Parameters.Add(new SqlParameter("AgentTimeE", m_OwnerMaintain.AgentTimeE));
                    m_cmd.Parameters.Add(new SqlParameter("ModifyUser", p_EmployeeEntity.WorkID));

                    m_cmd.Parameters.Add(new SqlParameter("OwnerType", "ISSUE"));

                    m_cmd.Parameters.Add(new SqlParameter("ParamValue1", m_OwnerMaintain.Site));
                    m_cmd.Parameters.Add(new SqlParameter("ParamValue2", p_dicOwnerMaintain["System1"].ToString()));
                    m_cmd.Parameters.Add(new SqlParameter("ParamValue3", p_dicOwnerMaintain["System3"].ToString()));

                    m_cmd.CommandText = m_SQL;
                    m_TransResult = g_DC.ExecuteNonQuery(p_EmployeeEntity, m_cmd);


                }
            }
            catch (Exception ex)
            {
                m_TransResult.LogMessage = ex.Message;
                m_TransResult.isSuccess = false;
            }

            return m_TransResult;

        }


        #endregion


        #region 多國語言

        public SysEntity.TransResult GetMultiLanguage(SysEntity.Employee p_EmployeeEntity, SysEntity.MultiLanguage p_MultiLanguage)
        {
            SysEntity.TransResult m_TransResult = new SysEntity.TransResult();
            try
            {
                string m_SQL = "SELECT [LanguageKey],[ValueTW],[ValueCN] ,[ValueUS] ,[ValueJP] ,[ValueBR]  FROM [dbo].[MultiLanguage]   ";
                using (SqlCommand m_cmd = g_DC.GetQueryCommand(p_EmployeeEntity, m_SQL, p_MultiLanguage))
                {
                    m_TransResult = g_DC.GetDataTable(p_EmployeeEntity, m_cmd);
                }
            }
            catch (Exception ex)
            {
                m_TransResult.LogMessage = ex.Message;
                m_TransResult.isSuccess = false;
            }
            return m_TransResult;
        }

        public SysEntity.TransResult GetListMultiLanguage(SysEntity.Employee p_EmployeeEntity, string p_LANGUAGE, string[] p_Keys)
        {
            SysEntity.TransResult m_TransResult = new SysEntity.TransResult();
            try
            {
                p_Keys = p_Keys.Distinct().ToArray();
                Dictionary<string, Object> m_Return = new Dictionary<string, object>();
                string m_SQL = "SELECT [ValueTW],[ValueCN],[ValueUS],[ValueJP],[ValueBR],[LanguageKey] FROM [dbo].[MultiLanguage] where [LanguageKey] in(  ";
                Int32 KeyCount = 1;
                foreach (string m_key in p_Keys)
                {
                    m_SQL += ((KeyCount == 1) ? "" : ",") + "@LanguageKey" + KeyCount.ToString();
                    KeyCount++;
                }
                m_SQL += ")";

                DataTable m_dtLan = new DataTable("Table");
                KeyCount = 1;
                using (SqlCommand m_cmd = new SqlCommand(m_SQL))
                {
                    foreach (string m_key in p_Keys)
                    {
                        m_cmd.Parameters.Add(new SqlParameter("LanguageKey" + KeyCount.ToString(), m_key.Replace("thlbl", "").Replace("lbl", "").Replace("Qbtn", "").Replace("btn", "").ToUpper()));
                        KeyCount++;
                    }
                    SysEntity.TransResult m_LangResult = g_DC.GetDataTable(p_EmployeeEntity, m_cmd);

                    if (m_LangResult.isSuccess)
                    {
                        m_dtLan = (DataTable)m_LangResult.ResultEntity;
                    }
                }
                foreach (string m_key in p_Keys)
                {
                    string m_LanguageKey = m_key.Replace("thlbl", "").Replace("lbl", "").Replace("Qbtn", "").Replace("btn", "");
                    if (m_dtLan.Rows.Count == 0)
                    {
                        m_Return[m_key] = m_LanguageKey;
                    }
                    else
                    {
                        string m_LanguageValue = "";
                        if (m_key.IndexOf("titleModuleName") == -1)
                        {
                            foreach (DataRow dr in m_dtLan.Select("LanguageKey='" + m_LanguageKey + "'"))
                            {
                                m_LanguageValue = dr["Value" + p_LANGUAGE].ToString();
                            }
                        }
                        else
                        {
                            Dictionary<string, Object> m_ModuleInf = new Dictionary<string, object>();
                            m_ModuleInf["ModuleKey"] = m_key.Replace("titleModuleName", "");
                            SysEntity.TransResult m_TransResultModuleInfo = GetModuleInfos(p_EmployeeEntity, m_ModuleInf);
                            if (m_TransResultModuleInfo.isSuccess)
                            {
                                foreach (DataRow dr in ((DataTable)m_TransResultModuleInfo.ResultEntity).Rows)
                                {
                                    if (p_LANGUAGE == "TW")
                                    {
                                        m_LanguageValue = dr["ModuleNameTW"].ToString();
                                    }
                                    else
                                    {
                                        m_LanguageValue = dr["ModuleNameUS"].ToString();
                                    }
                                }


                            }

                        }


                        if (m_LanguageValue == "")
                        {
                            m_Return[m_key] = m_LanguageKey;
                        }
                        else
                        {
                            m_Return[m_key] = m_LanguageValue;

                        }
                    }
                }


                m_TransResult.isSuccess = true;

                m_TransResult.ResultEntity = m_Return;
            }
            catch (Exception ex)
            {
                m_TransResult.isSuccess = false;
                m_TransResult.LogMessage = ex.Message;
            }
            return m_TransResult;
        }
        #endregion


        #region 權限


        public DataTable GetAuthIDByWorkID(SysEntity.Employee p_EmployeeEntity, string p_WorkID)
        {

            DataTable m_dt = new DataTable();
            string m_SQL = "SELECT [AuthID] ,[referID] FROM [dbo].[AuthAssign] where [AssignType]='Employee' and [referID]=@WorkID ";
            m_SQL += " union all SELECT [AuthID] ,[referID] FROM [dbo].[AuthAssign] where [AssignType]='Group' and ";
            m_SQL += " [referID] in(select GroupID from  [dbo].[EmployeeGroup] where WorkID=@WorkID) ";

            using (SqlCommand m_cmd = new SqlCommand(m_SQL))
            {
                m_cmd.Parameters.Add(new SqlParameter("WorkID", p_WorkID));

                SysEntity.TransResult m_TransResult = new SysEntity.TransResult();


                m_TransResult = g_DC.GetDataTable(p_EmployeeEntity, m_cmd);
                if (m_TransResult.isSuccess)
                {
                    m_dt = (DataTable)m_TransResult.ResultEntity;
                }
            }
            return m_dt;
        }

        public Boolean CheckIsManager(SysEntity.Employee p_EmployeeEntity)
        {
            Boolean m_isManager = false;
            DataTable m_dt = new DataTable();
            string m_SQL = "SELECT PERNR FROM  [dbo].[DEPARTMENTINFO] WHERE PERNR=@PERNR ";
            using (SqlCommand m_cmd = new SqlCommand(m_SQL))
            {
                m_cmd.Parameters.Add(new SqlParameter("PERNR", p_EmployeeEntity.WorkID));
                SysEntity.TransResult m_TransResult = new SysEntity.TransResult();
                m_TransResult = g_DC.GetDataTable(p_EmployeeEntity, m_cmd);
                if (m_TransResult.isSuccess)
                {
                    foreach (DataRow dr in ((DataTable)m_TransResult.ResultEntity).Rows)
                    {
                        m_isManager = true;
                    }
                }
            }
            return m_isManager;
        }

        public DataTable GetAuthPERNRByWorkID(SysEntity.Employee p_EmployeeEntity, string p_WorkID)
        {
            DataTable m_dt = new DataTable();
            string m_SQL = " SELECT PERNR FROM [dbo].[EMPLOYEEINFO] WHERE ORGEH IN  ";
            m_SQL += " (SELECT DISTINCT [DEPTID] FROM [dbo].[DEPTORGANIZATION] O ";
            m_SQL += " LEFT JOIN [dbo].[DEPARTMENTINFO] D ON O.DEPTID=D.OBJID ";
            m_SQL += " LEFT JOIN (SELECT  OBJID,OBJIDU,SHORT1,D.PERNR,E.ENAME FROM [dbo].[DEPARTMENTINFO] D  ";
            m_SQL += " LEFT JOIN [dbo].[EMPLOYEEINFO] E ON D.PERNR=E.PERNR ) D1 ON O.DEPTOWNER=D1.OBJID  ";
            m_SQL += " WHERE D1.PERNR=@PERNR  ";
            m_SQL += " UNION ALL SELECT OBJID [DEPTID] FROM [dbo].[DEPARTMENTINFO] WHERE PERNR=@PERNR) ";

            using (SqlCommand m_cmd = new SqlCommand(m_SQL))
            {
                m_cmd.Parameters.Add(new SqlParameter("PERNR", p_WorkID));
                SysEntity.TransResult m_TransResult = new SysEntity.TransResult();
                m_TransResult = g_DC.GetDataTable(p_EmployeeEntity, m_cmd);
                if (m_TransResult.isSuccess)
                {
                    m_dt = (DataTable)m_TransResult.ResultEntity;
                }
            }
            return m_dt;
        }

        public DataTable GetSiteFromStatusBySiteFromID(SysEntity.Employee p_EmployeeEntity, string p_SiteFromID, string p_FieldID)
        {

            DataTable m_dt = new DataTable();
            string m_SQL = "SELECT S.[SiteFormID],[StatusTitle],[StatusKey],[ConditionType],[StatusCondition],[ConditionRef],v.EventID,v.EventActionType FROM [dbo].[SiteFormStatus] S left join  (select SiteFormID,FieldID,V.EventID,E.EventActionType from [dbo].ViewDetail V left join  [dbo].EventSetting E on V.EventID=E.EventID ) V on S.SiteFormID=V.SiteFormID and S.FieldID=V.FieldID where S.SiteFormID=@SiteFormID and S.FieldID=@FieldID Order by OrderIndex ";

            using (SqlCommand m_cmd = new SqlCommand(m_SQL))
            {
                m_cmd.Parameters.Add(new SqlParameter("SiteFormID", p_SiteFromID));
                m_cmd.Parameters.Add(new SqlParameter("FieldID", p_FieldID));
                SysEntity.TransResult m_TransResult = new SysEntity.TransResult();
                m_TransResult = g_DC.GetDataTable(p_EmployeeEntity, m_cmd);
                if (m_TransResult.isSuccess)
                {
                    m_dt = (DataTable)m_TransResult.ResultEntity;
                }
            }
            return m_dt;
        }


        public DataTable GetSiteFromStatusFieldIDsBySiteFromID(SysEntity.Employee p_EmployeeEntity, string p_SiteFromID)
        {

            DataTable m_dt = new DataTable();
            string m_SQL = "SELECT distinct S.FieldID FROM [dbo].[SiteFormStatus] S left join  (select SiteFormID,FieldID,EventID from [dbo].ViewDetail) V on S.SiteFormID=V.SiteFormID and  S.FieldID=V.FieldID  where S.SiteFormID=@SiteFormID   order by FieldID ";

            using (SqlCommand m_cmd = new SqlCommand(m_SQL))
            {
                m_cmd.Parameters.Add(new SqlParameter("SiteFormID", p_SiteFromID));

                SysEntity.TransResult m_TransResult = new SysEntity.TransResult();


                m_TransResult = g_DC.GetDataTable(p_EmployeeEntity, m_cmd);
                if (m_TransResult.isSuccess)
                {
                    m_dt = (DataTable)m_TransResult.ResultEntity;
                }
            }
            return m_dt;
        }


        public bool GetSignerButtonByDocID(SysEntity.Employee p_EmployeeEntity, string p_DocID)
        {
            bool m_isSign = true;
            DataTable m_dt = new DataTable();



            string m_SQL = "SELECT [DocID] FROM [dbo].[SignRecord] where [DocID]=@DocID   ";

            using (SqlCommand m_cmd = new SqlCommand(m_SQL))
            {
                m_cmd.Parameters.Add(new SqlParameter("DocID", p_DocID));
                SysEntity.TransResult m_TransResult = new SysEntity.TransResult();
                m_TransResult = g_DC.GetDataTable(p_EmployeeEntity, m_cmd);
                if (m_TransResult.isSuccess)
                {
                    m_dt = (DataTable)m_TransResult.ResultEntity;
                    if (m_dt.Rows.Count == 0)
                    {
                        m_isSign = true;
                    }
                    else
                    {
                        m_isSign = false;
                    }
                }
            }

            if (!(m_isSign))
            {
                m_SQL = "SELECT [Signer] FROM [dbo].[DocSigning] where [DocID]=@DocID and Signer=@Signer  ";

                using (SqlCommand m_cmd = new SqlCommand(m_SQL))
                {
                    m_cmd.Parameters.Add(new SqlParameter("DocID", p_DocID));
                    m_cmd.Parameters.Add(new SqlParameter("Signer", p_EmployeeEntity.WorkID));

                    SysEntity.TransResult m_TransResult = new SysEntity.TransResult();

                    m_TransResult = g_DC.GetDataTable(p_EmployeeEntity, m_cmd);
                    if (m_TransResult.isSuccess)
                    {
                        m_dt = (DataTable)m_TransResult.ResultEntity;
                        if (m_dt.Rows.Count == 0)
                        {
                            m_isSign = false;
                        }
                        else
                        {
                            m_isSign = true;
                        }
                    }
                }
            }

            return m_isSign;
        }


        public bool CheckIsSigner(SysEntity.Employee p_EmployeeEntity, string p_DocID)
        {
            bool m_isSigner = true;
            DataTable m_dt = new DataTable();
            string m_SQL = "SELECT [DocID] FROM [dbo].[DocSigning] where [DocID]=@DocID and [Signer]=@Signer   ";
            using (SqlCommand m_cmd = new SqlCommand(m_SQL))
            {
                m_cmd.Parameters.Add(new SqlParameter("DocID", p_DocID));
                m_cmd.Parameters.Add(new SqlParameter("Signer", p_EmployeeEntity.WorkID));

                SysEntity.TransResult m_TransResult = new SysEntity.TransResult();
                m_TransResult = g_DC.GetDataTable(p_EmployeeEntity, m_cmd);
                if (m_TransResult.isSuccess)
                {
                    m_dt = (DataTable)m_TransResult.ResultEntity;
                    if (m_dt.Rows.Count == 0)
                    {
                        m_isSigner = false;
                    }
                }
            }
            return m_isSigner;
        }


        #endregion

        #region 轉送

        public SysEntity.TransResult CommitFormTransfer(SysEntity.Employee p_EmployeeEntity, Dictionary<string, Object> p_DicFormTransfer, List<Business_Logic.Entity.SysEntity.CommandEntity> p_CommandEntitys)
        {
            SysEntity.TransResult m_TransResult = new SysEntity.TransResult();

            SysEntity.CommandEntity m_CommandEntity = new SysEntity.CommandEntity();
            try
            {
                if (p_DicFormTransfer.ContainsKey("TransferRemark"))
                {
                    Dictionary<string, Object> m_disFormTransferLog = new Dictionary<string, object>();
                    m_disFormTransferLog["FormTransferID"] = DateTime.Now.ToString("yyyyMMddhhmmssfff");
                    m_disFormTransferLog["TransferUserFrom"] = p_EmployeeEntity.WorkID;
                    m_disFormTransferLog["TransferUserTo"] = p_DicFormTransfer["WorkID"];
                    m_disFormTransferLog["TransferRemark"] = p_DicFormTransfer["TransferRemark"];
                    m_disFormTransferLog["TransferKey"] = p_DicFormTransfer["TransferKey"];
                    m_disFormTransferLog["CreateUser"] = p_EmployeeEntity.WorkID;

                    m_CommandEntity = g_DC.GetInsertCommand(p_EmployeeEntity, "FormTransferLog", m_disFormTransferLog);
                    p_CommandEntitys.Add(m_CommandEntity);
                    m_TransResult = g_DC.ExecuteCommandEntitys(p_EmployeeEntity, p_CommandEntitys);

                }
                m_TransResult.isSuccess = true;

            }
            catch (Exception ex)
            {
                m_TransResult.LogMessage = ex.Message;
                m_TransResult.isSuccess = false;
            }
            return m_TransResult;
        }

        public SysEntity.TransResult FormTransferSample(SysEntity.Employee p_EmployeeEntity, Dictionary<string, Object> p_DicFormTransfer)
        {
            SysEntity.TransResult m_TransResult = new SysEntity.TransResult();
            try
            {
                // p_DicFormTransfer["WorkID"].ToString(); //轉送給誰
                // p_DicFormTransfer["TransferRemark"].ToString();//轉送備註
                // p_DicFormTransfer["FromKeys"].ToString(); //轉送Key


                // 多筆 SQL Command
                List<Business_Logic.Entity.SysEntity.CommandEntity> m_MultiCommandEntitys = new List<SysEntity.CommandEntity>();

                m_TransResult = g_FH.CheckFormKeys(p_DicFormTransfer);
                if (m_TransResult.isSuccess)
                {
                    Dictionary<string, Object> m_Params = (Dictionary<string, Object>)m_TransResult.ResultEntity;
                    List<object> m_lisParams = new List<object>();
                    //SQL Command 1
                    Business_Logic.Entity.SysEntity.CommandEntity m_MultiCommandEntity = new SysEntity.CommandEntity();
                    object m_Param = new
                    {
                        FormID = m_Params["FormID"].ToString(),
                        Description = m_Params["Description"].ToString()
                    };
                    m_MultiCommandEntity.SQLCommand = "update [dbo].[Form] set Description=@Description  Where [FormID]=@FormID";
                    m_MultiCommandEntity.Parameters = m_Param;
                    m_MultiCommandEntitys.Add(m_MultiCommandEntity);

                    //SQL Command 2
                    m_MultiCommandEntity = new SysEntity.CommandEntity();
                    m_MultiCommandEntity.SQLCommand = "update [dbo].[SiteForm] set AgentName=@AgentName  Where [FormID]=@FormID";

                    m_Param = new
                    {
                        FormID = m_Params["FormID"].ToString(),
                        AgentName = m_Params["Description"].ToString()
                    };
                    m_MultiCommandEntity.Parameters = m_Param;
                    m_MultiCommandEntitys.Add(m_MultiCommandEntity);

                    p_DicFormTransfer["TransferKey"] = m_Params["FormID"];
                    m_TransResult = CommitFormTransfer(p_EmployeeEntity, p_DicFormTransfer, m_MultiCommandEntitys);

                }



            }
            catch (Exception ex)
            {
                m_TransResult.LogMessage = ex.Message;
                m_TransResult.isSuccess = false;
            }
            return m_TransResult;
        }

        public SysEntity.TransResult RejectSample(SysEntity.Employee p_EmployeeEntity, Dictionary<string, Object> p_DicReject)
        {
            SysEntity.TransResult m_TransResult = new SysEntity.TransResult();
            try
            {
                //p_DicFormTransfer["RejectRemark"].ToString();//駁回備註
                //p_DicFormTransfer["FromKeys"].ToString(); //駁回Key


                // 多筆 SQL Command
                List<Business_Logic.Entity.SysEntity.CommandEntity> m_MultiCommandEntitys = new List<SysEntity.CommandEntity>();

                m_TransResult = g_FH.CheckFormKeys(p_DicReject);
                if (m_TransResult.isSuccess)
                {
                    Dictionary<string, Object> m_Params = (Dictionary<string, Object>)m_TransResult.ResultEntity;
                    List<object> m_lisParams = new List<object>();
                    //SQL Command 1
                    Business_Logic.Entity.SysEntity.CommandEntity m_MultiCommandEntity = new SysEntity.CommandEntity();
                    object m_Param = new
                    {
                        FormID = m_Params["FormID"].ToString(),
                        Description = m_Params["Description"].ToString()
                    };
                    m_MultiCommandEntity.SQLCommand = "update [dbo].[Form] set Description=@Description  Where [FormID]=@FormID";
                    m_MultiCommandEntity.Parameters = m_Param;
                    m_MultiCommandEntitys.Add(m_MultiCommandEntity);

                    //SQL Command 2
                    m_MultiCommandEntity = new SysEntity.CommandEntity();
                    m_MultiCommandEntity.SQLCommand = "update [dbo].[SiteForm] set AgentName=@AgentName  Where [FormID]=@FormID";

                    m_Param = new
                    {
                        FormID = m_Params["FormID"].ToString(),
                        AgentName = m_Params["Description"].ToString()
                    };
                    m_MultiCommandEntity.Parameters = m_Param;
                    m_MultiCommandEntitys.Add(m_MultiCommandEntity);

                    p_DicReject["TransferKey"] = m_Params["FormID"];
                    m_TransResult = CommitFormTransfer(p_EmployeeEntity, p_DicReject, m_MultiCommandEntitys);

                }



            }
            catch (Exception ex)
            {
                m_TransResult.LogMessage = ex.Message;
                m_TransResult.isSuccess = false;
            }
            return m_TransResult;
        }


        #endregion

        #region 批簽
        public SysEntity.TransResult BatchApproveSample(SysEntity.Employee p_EmployeeEntity, Dictionary<string, Object> p_BatchApprove)
        {
            SysEntity.TransResult m_TransResult = new SysEntity.TransResult();
            try
            {
                string m_BatchApproveRemark = p_BatchApprove["BatchApproveRemark"].ToString();//批簽備註


                // 多筆 SQL Command
                List<Business_Logic.Entity.SysEntity.CommandEntity> m_MultiCommandEntity = new List<SysEntity.CommandEntity>();

                m_TransResult = g_FH.CheckFormKeys(p_BatchApprove);
                if (m_TransResult.isSuccess)
                {

                    List<Dictionary<string, Object>> m_FormKeys = (List<Dictionary<string, Object>>)m_TransResult.ResultEntity;
                    Business_Logic.Entity.SysEntity.CommandEntity m_CommandEntity = new SysEntity.CommandEntity();


                    foreach (Dictionary<string, Object> m_FormKey in m_FormKeys)
                    {
                        m_CommandEntity = new SysEntity.CommandEntity();
                        object m_Param = new
                        {
                            FormID = m_FormKey["FormID"].ToString(),
                         
                            ModifyUser = p_EmployeeEntity.WorkID,
                            Description = m_BatchApproveRemark
                        };
                        m_CommandEntity.SQLCommand = "update [dbo].[Form] set Description=@Description , ModifyUser=@ModifyUser , ModifyDate=(Convert(varchar(10),Getdate(),111)), ModifyTime=(Convert(varchar(8),Getdate(),108))  Where [FormID]=@FormID  ";
                        m_CommandEntity.Parameters = m_Param;
                        m_MultiCommandEntity.Add(m_CommandEntity);
                    }
                    m_TransResult = g_DC.ExecuteCommandEntitys(p_EmployeeEntity, m_MultiCommandEntity);

                }
            }
            catch (Exception ex)
            {
                m_TransResult.LogMessage = ex.Message;
                m_TransResult.isSuccess = false;
            }
            return m_TransResult;
        }

        #endregion

        #region PageCarousel


        public SysEntity.TransResult GetPeriodCarousel(SysEntity.Employee p_EmployeeEntity, string p_SiteID)
        {
            SysEntity.TransResult m_TransResult = new SysEntity.TransResult();

            try
            {
                string m_SQL = "  SELECT P.*,F.FileKey,F.FileEntity FROM [dbo].[PageCarousel] P Left Join ( select DocID,FileKey,FileEntity from  [dbo].[FileFolder]) F " +
                    " on P.[DocID]=F.[DocID] where [SiteID]=@SiteID and  (CONVERT([varchar](10),getdate(),(111))) " +
                    "between [Start_Period] and [End_Period] order by [Start_Period] ";

                using (SqlCommand m_cmd = new SqlCommand(m_SQL))
                {
                    m_cmd.Parameters.Add(new SqlParameter("SiteID", p_SiteID));
                    m_TransResult = g_DC.GetDataTable(p_EmployeeEntity, m_cmd);
                }
            }
            catch (Exception ex)
            {
                m_TransResult.isSuccess = false;
                m_TransResult.LogMessage = ex.Message;
            }


            return m_TransResult;
        }


        public SysEntity.TransResult GetPageCarousels(SysEntity.Employee p_EmployeeEntity, Dictionary<string, Object> p_dicPageCarousel)
        {
            SysEntity.TransResult m_TransResult = new SysEntity.TransResult();
            Dictionary<string, Object> m_objBetween = new Dictionary<string, Object>();
            SysEntity.PageCarousel m_PageCarousel = g_FH.ConvertEntity<SysEntity.PageCarousel>(p_dicPageCarousel, ref m_objBetween);

            try
            {
                string m_SQL = "SELECT * FROM [dbo].[PageCarousel]   ";
                using (SqlCommand m_cmd = g_DC.GetQueryCommand(p_EmployeeEntity, m_SQL, m_PageCarousel, m_objBetween))
                {
                    m_TransResult = g_DC.GetDataTable(p_EmployeeEntity, m_cmd);
                }
            }
            catch (Exception ex)
            {
                m_TransResult.LogMessage = ex.Message;
                m_TransResult.isSuccess = false;
            }
            return m_TransResult;
        }

        public SysEntity.TransResult GetPageCarouselGrid(SysEntity.Employee p_EmployeeEntity, Dictionary<string, Object> p_dicPageCarousel, string p_PageIndex, string p_PageSize, Dictionary<string, Object> p_Param)
        {

            SysEntity.TransResult m_TransResult = new SysEntity.TransResult();
            Dictionary<string, Object> m_objBetween = new Dictionary<string, Object>();
            SysEntity.PageCarousel m_PageCarousel = g_FH.ConvertEntity<SysEntity.PageCarousel>(p_dicPageCarousel, ref m_objBetween);
            ArrayList m_alGrid = new ArrayList();
            string[] m_LikeProp = { "CarouselDepiction", "CarouselType" };

            try
            {
                string m_SQL = "SELECT [DocID],[SiteID],[CarouselType],[CarouselCaption],[CarouselDepiction],[Start_Period],[End_Period],[Status] FROM [dbo].[PageCarousel]    ";
                using (SqlCommand m_cmd = g_DC.GetQueryCommand(p_EmployeeEntity, m_SQL, m_PageCarousel, m_objBetween, m_LikeProp))
                {

                    m_TransResult = g_DC.GetDataTableByGrid(p_EmployeeEntity, m_cmd, p_Param, p_PageSize, p_PageIndex);
                    if (m_TransResult.isSuccess)
                    {
                        m_TransResult = g_DC.GridDataCtrl(m_TransResult, p_PageSize, p_Param);
                    }
                }
            }
            catch (Exception ex)
            {
                m_TransResult.LogMessage = ex.Message;
                m_TransResult.isSuccess = false;
            }
            return m_TransResult;
        }

        public SysEntity.TransResult AddPageCarousel(SysEntity.Employee p_EmployeeEntity, Dictionary<string, Object> p_dicPageCarousel)
        {
            SysEntity.TransResult m_TransResult = new SysEntity.TransResult();
            try
            {
                m_TransResult = g_DC.InsertIntoByEntity<SysEntity.PageCarousel>(p_EmployeeEntity, "PageCarousel", p_dicPageCarousel);
            }
            catch (Exception ex)
            {
                m_TransResult.LogMessage = ex.Message;
                m_TransResult.isSuccess = false;
            }

            return m_TransResult;

        }

        public SysEntity.TransResult DelPageCarousel(SysEntity.Employee p_EmployeeEntity, Dictionary<string, Object> p_dicPageCarousel)
        {
            SysEntity.TransResult m_TransResult = new SysEntity.TransResult();
            Dictionary<string, Object> m_ArrayList = new Dictionary<string, Object>();

            SysEntity.PageCarousel m_PageCarousel = g_FH.ConvertEntity<SysEntity.PageCarousel>(p_dicPageCarousel, ref m_ArrayList);

            try
            {
                string m_SQL = "Delete [dbo].[PageCarousel] Where [DocID]=@DocID";
                using (SqlCommand m_cmd = new SqlCommand(m_SQL))
                {
                    m_cmd.Parameters.Add(new SqlParameter("PageCarouselID", m_PageCarousel.DocID));
                    m_TransResult = g_DC.ExecuteNonQuery(p_EmployeeEntity, m_cmd);
                }
            }
            catch (Exception ex)
            {
                m_TransResult.LogMessage = ex.Message;
                m_TransResult.isSuccess = false;
            }

            return m_TransResult;

        }

        public SysEntity.TransResult UpdPageCarousel(SysEntity.Employee p_EmployeeEntity, Dictionary<string, Object> p_dicPageCarousel)
        {
            SysEntity.TransResult m_TransResult = new SysEntity.TransResult();
            Dictionary<string, Object> m_ArrayList = new Dictionary<string, Object>();

            SysEntity.PageCarousel m_PageCarousel = g_FH.ConvertEntity<SysEntity.PageCarousel>(p_dicPageCarousel, ref m_ArrayList);

            try
            {
                string m_SQL = "";
                using (SqlCommand m_cmd = new SqlCommand())
                {
                    m_SQL = "Update [dbo].[PageCarousel]  set  CarouselType =@CarouselType ,SiteID =@SiteID , Status =@Status , End_Period =@End_Period ,Start_Period =@Start_Period , CarouselCaption =@CarouselCaption , CarouselDepiction =@CarouselDepiction,CarouselLink =@CarouselLink,ModifyUser=@ModifyUser , ModifyDate=(Convert(varchar(10),Getdate(),111)), ModifyTime=(Convert(varchar(8),Getdate(),108)) ";
                    m_cmd.Parameters.Add(new SqlParameter("CarouselCaption", m_PageCarousel.CarouselCaption));
                    m_cmd.Parameters.Add(new SqlParameter("CarouselDepiction", m_PageCarousel.CarouselDepiction));
                    m_cmd.Parameters.Add(new SqlParameter("CarouselLink", m_PageCarousel.CarouselLink));
                    m_cmd.Parameters.Add(new SqlParameter("Start_Period", m_PageCarousel.Start_Period));
                    m_cmd.Parameters.Add(new SqlParameter("End_Period", m_PageCarousel.End_Period));
                    m_cmd.Parameters.Add(new SqlParameter("Status", m_PageCarousel.Status));
                    m_cmd.Parameters.Add(new SqlParameter("CarouselType", m_PageCarousel.CarouselType));

                    m_cmd.Parameters.Add(new SqlParameter("ModifyUser", p_EmployeeEntity.WorkID));

                    if (m_PageCarousel.DocID != "")
                    {
                        m_SQL += " where [DocID]=@DocID and SiteID=@SiteID  ";
                        m_cmd.Parameters.Add(new SqlParameter("DocID", m_PageCarousel.DocID));
                        m_cmd.Parameters.Add(new SqlParameter("SiteID", m_PageCarousel.SiteID));

                        m_cmd.CommandText = m_SQL;
                        m_TransResult = g_DC.ExecuteNonQuery(p_EmployeeEntity, m_cmd);
                    }
                    else
                    {
                        m_TransResult.LogMessage = "DocID 不可為空!!!";
                        m_TransResult.isSuccess = false;
                    }

                }
            }
            catch (Exception ex)
            {
                m_TransResult.LogMessage = ex.Message;
                m_TransResult.isSuccess = false;
            }

            return m_TransResult;

        }
        #endregion

        #region Q&A Methods From Danny_Chu

        public SysEntity.TransResult GetQAMaster(SysEntity.Employee p_EmployeeEntity, Dictionary<string, Object> p_dicQAMaster)
        {
            SysEntity.TransResult m_TransResult = new SysEntity.TransResult();
            Dictionary<string, Object> m_objBetween = new Dictionary<string, Object>();
            SysEntity.QAMaster m_QAMaster = g_FH.ConvertEntity<SysEntity.QAMaster>(p_dicQAMaster, ref m_objBetween);

            try
            {
                string m_SQL = "SELECT [QAID],[DocID]," +
                    "dbo.fun_GetEmployeeNameExt(Case_Owner) as Case_OwnerALL," +
                    "dbo.fun_GetEmployeeNameExt(ModifyUser) as ModifyUserALL," +
                    "[QASubject],[QADescription],[QAReply],[Status],[Priority],[System1],[System3],[Issue_Type],[User_Department]," +
                    "[Company],[Site],[KeyWords],Finish_Date+' '+ Finish_Time as Finish_Date2,[CreateUser],[CreateDate],ModifyUser, ModifyDate+' '+ModifyTime as ModifyDate2 ,User_Contact,Case_Owner,CreateTime,ModifyTime " +
                    "FROM [dbo].[QAMaster]   ";
                using (SqlCommand m_cmd = g_DC.GetQueryCommand(p_EmployeeEntity, m_SQL, m_QAMaster, m_objBetween))
                {
                    m_TransResult = g_DC.GetDataTable(p_EmployeeEntity, m_cmd);
                    //帶去給轉派用
                    //p_EmployeeEntity.User_ContactALL = m_QAMaster.User_ContactALL;
                }
            }
            catch (Exception ex)
            {
                m_TransResult.LogMessage = ex.Message;
                m_TransResult.isSuccess = false;
            }
            return m_TransResult;
        }

        public SysEntity.TransResult GetQAMasterGrid(SysEntity.Employee p_EmployeeEntity, Dictionary<string, Object> p_dicQAMaster, string p_PageIndex, string p_PageSize, Dictionary<string, Object> p_Param)
        {
            SysEntity.TransResult m_TransResult = new SysEntity.TransResult();
            Dictionary<string, Object> m_objBetween = new Dictionary<string, Object>();
            m_objBetween["CreateDate_S"] = p_dicQAMaster["CreateDate_S"].ToString();
            m_objBetween["CreateDate_E"] = p_dicQAMaster["CreateDate_E"].ToString();

            //SysEntity.QAMaster m_QAMaster = g_FunctionHandler.ConvertEntity<SysEntity.QAMaster>(p_dicQAMaster, ref m_objBetween);
            ArrayList m_alGrid = new ArrayList();

            p_dicQAMaster["CustKeyWords"] = p_dicQAMaster["KeyWords"].ToString();
            p_dicQAMaster["KeyWords"] = "";
            string[] m_LikeProp = { "QAID", "CustKeyWords", "User_Contact_ALL", "Case_Owner_ALL" };

            try
            {
               
               
                string m_SQL = "" +
                    " SELECT Q.QAID, Q.QASubject, Q.QADescription, Q.QAReply, " +
                    "        MP3.sText" + p_EmployeeEntity.CurrentLanguage + " AS Issue_Type_OK, " +
                    "        MP1.sText" + p_EmployeeEntity.CurrentLanguage + " AS Status_OK," +
                    "        ISNULL(dbo.fun_GetEmployeeName(Q.User_Contact),Q.User_Contact) AS User_Contact_ALL, " +
                    "        ISNULL(dbo.fun_GetEmployeeName(Q.Case_Owner),Q.Case_Owner) AS Case_Owner_ALL, Q.Priority, " +
                    "        Q.System1, Q.System3, Q.Issue_Type, Q.KeyWords, " +
                    "        Q.ModifyUser, " +
                    "        Q.QASubject + Q.QADescription + Q.QAReply + Q.KeyWords AS CustKeyWords," +
                    "        Q.Case_Owner, Q.User_Contact,  " +
                    "        MP2.sText" + p_EmployeeEntity.CurrentLanguage + " AS Priority_OK, " +
                    "        MP4.sText" + p_EmployeeEntity.CurrentLanguage + " AS System1_OK, " +
                    "        MP.sText" + p_EmployeeEntity.CurrentLanguage + " AS System3_OK, " +
                    "        Q.Status,Q.CreateDate+' '+Q.CreateTime AS CreateDateTime,Q.CreateDate ,Site,User_Department,Company" +
                    "  FROM dbo.QAMaster Q  " +
                    "  LEFT OUTER JOIN dbo.MultiParameter MP ON Q.System3 = MP.sValue  " +
                    "  LEFT OUTER JOIN dbo.MultiParameter MP4 ON Q.System1 = MP4.sValue  " +
                    "  LEFT OUTER JOIN dbo.MultiParameter MP3 ON Q.Issue_Type = MP3.sValue  " +
                    "  LEFT OUTER JOIN dbo.MultiParameter MP2 ON Q.Priority = MP2.sValue  " +
                    "  LEFT OUTER JOIN dbo.MultiParameter MP1 ON Q.Status = MP1.sValue ";

                using (SqlCommand m_cmd = g_DC.GetQueryCommand(p_EmployeeEntity, m_SQL, p_dicQAMaster, m_objBetween, m_LikeProp))
                {
                    m_TransResult = g_DC.GetDataTableByGrid(p_EmployeeEntity, m_cmd, p_Param, p_PageSize, p_PageIndex);
                    if (m_TransResult.isSuccess)
                    {
                        m_TransResult = g_DC.GridDataCtrl(m_TransResult, p_PageSize, p_Param);
                    }
                }
            }
            catch (Exception ex)
            {
                m_TransResult.LogMessage = ex.Message;
                m_TransResult.isSuccess = false;
            }

            return m_TransResult;
        }

        public SysEntity.TransResult AddQAMaster(SysEntity.Employee p_EmployeeEntity, Dictionary<string, Object> p_dicQAMaster)
        {
            SysEntity.TransResult m_TransResult = new SysEntity.TransResult();
            Dictionary<string, Object> m_ArrayList = new Dictionary<string, Object>();
            SysEntity.QAMaster m_QAMaster = g_FH.ConvertEntity<SysEntity.QAMaster>(p_dicQAMaster, ref m_ArrayList);
            try
            {
                SysEntity.TransResult m_TransResultObj = new SysEntity.TransResult();
               
             
                SysEntity.Site m_TransResultSite = new SysEntity.Site();

                

                using (SqlCommand m_cmd = new SqlCommand())
                {
                    //From 掛號程式 For QAID
                    string sNewQAID = RegisterQAID();

                    string System1 = "", System3 = "", sCase_Owner = "";
                    System1 = m_QAMaster.System1;
                    System3 = m_QAMaster.System3;
                  
                    Dictionary<string, Object> p_dicOwnerMaintain = new Dictionary<string, object>();

                    p_dicOwnerMaintain["System1"] = System1;
                    p_dicOwnerMaintain["System3"] = System3;


                    SysEntity.TransResult m_OwnerResult = GetOwnerMaintain(p_EmployeeEntity, p_dicOwnerMaintain);
                    if (m_OwnerResult.isSuccess)
                    {
                        foreach (DataRow dr in ((DataTable)m_OwnerResult.ResultEntity).Rows)
                        {
                            sCase_Owner = dr["CaseOwner"].ToString();
                        }
                    }

                    string m_SQL = "";
                    m_SQL = "Update [dbo].[QAMaster]  set " +
                        "DocID=@DocID,User_Contact =@User_Contact  , User_Department =@User_Department ,Company =@Company , Site =@Site ," +
                        "Status =@Status , Priority =@Priority ," +
                        "System1 =@System1 ,System3 =@System3 ," +
                        "QASubject =@QASubject ,QADescription =@QADescription ," +
                        "KeyWords =@KeyWords,Case_Owner =@Case_Owner,Issue_Type =@Issue_Type," +
                        "CreateUser=@CreateUser , CreateDate=(Convert(varchar(10),Getdate(),111)), CreateTime=(Convert(varchar(8),Getdate(),108)), " +
                        "ModifyUser=@ModifyUser , ModifyDate=(Convert(varchar(10),Getdate(),111)), ModifyTime=(Convert(varchar(8),Getdate(),108)) ";

                    m_cmd.Parameters.Add(new SqlParameter("DocID", m_QAMaster.DocID));
                    m_cmd.Parameters.Add(new SqlParameter("User_Contact", m_QAMaster.User_Contact));
                    m_cmd.Parameters.Add(new SqlParameter("User_Department", m_QAMaster.User_Department));
                    m_cmd.Parameters.Add(new SqlParameter("CreateUser", m_QAMaster.User_Contact));
                    m_cmd.Parameters.Add(new SqlParameter("ModifyUser", m_QAMaster.User_Contact));

                    //日後有其他系統的ID進來  才能抓到其他的Company
                    m_cmd.Parameters.Add(new SqlParameter("Company", "ADATA"));
                    m_cmd.Parameters.Add(new SqlParameter("Site", m_TransResultSite.Description));

                    m_cmd.Parameters.Add(new SqlParameter("Status", m_QAMaster.Status));
                    m_cmd.Parameters.Add(new SqlParameter("Priority", m_QAMaster.Priority));
                    m_cmd.Parameters.Add(new SqlParameter("Issue_Type", m_QAMaster.Issue_Type));

                    m_cmd.Parameters.Add(new SqlParameter("System1", System1));
                    m_cmd.Parameters.Add(new SqlParameter("System3", System3));
                    m_cmd.Parameters.Add(new SqlParameter("Case_Owner", sCase_Owner));

                    m_cmd.Parameters.Add(new SqlParameter("QASubject", m_QAMaster.QASubject));
                    m_cmd.Parameters.Add(new SqlParameter("QADescription", m_QAMaster.QADescription));
                    m_cmd.Parameters.Add(new SqlParameter("KeyWords", m_QAMaster.KeyWords));
                    m_cmd.Parameters.Add(new SqlParameter("Finish_Date", ""));

                    if (sNewQAID != "")
                    {
                        m_SQL += " where [QAID]=@QAID ";
                        m_cmd.Parameters.Add(new SqlParameter("QAID", sNewQAID));
                        m_cmd.CommandText = m_SQL;
                        m_TransResult = g_DC.ExecuteNonQuery(p_EmployeeEntity, m_cmd);

                        if (m_TransResult.isSuccess)
                        {
                            #region Mail通知


                            #endregion
                        }

                    }
                    else
                    {
                        m_TransResult.LogMessage = "QAID can not be empty!!!";
                        m_TransResult.isSuccess = false;
                    }

                }
            }
            catch (Exception ex)
            {
                m_TransResult.LogMessage = ex.Message;
                m_TransResult.isSuccess = false;
            }
            return m_TransResult;
        }

        public string getEmail(SysEntity.Employee p_EmployeeEntity, string WorkIDs)
        {
            string[] FIPersonWorkID = WorkIDs.Split(',');
            string Email = "";
            string m_EmployeeID = "";
            try
            {
                foreach (string tmpWorkID in FIPersonWorkID)
                {
                    if (tmpWorkID.Length > 0)
                    {
                      
                    }
                }
            }
            catch
            {
                throw new Exception("Get Mail is fail Employee ID : " + m_EmployeeID + " is not exists !");
            }

            return Email;
        }

        public SysEntity.TransResult UpdQAMaster(SysEntity.Employee p_EmployeeEntity, Dictionary<string, Object> p_dicQAMaster)
        {
            SysEntity.TransResult m_TransResult = new SysEntity.TransResult();
            Dictionary<string, Object> m_ArrayList = new Dictionary<string, Object>();
            SysEntity.QAMaster m_QAMaster = g_FH.ConvertEntity<SysEntity.QAMaster>(p_dicQAMaster, ref m_ArrayList);

            m_TransResult.ResultEntity = "";
            if (m_QAMaster.QAReply == "" && m_QAMaster.Status == "273")
            {
                //m_TransResult.LogMessage = "請輸入 [答案] 才可結案!!";
                m_TransResult.LogMessage = "Please enter [Answer] to close the case!!";
                m_TransResult.isSuccess = false;
            }
            else
            {

                try
                {
                    string m_SQL = "";
                    using (SqlCommand m_cmd = new SqlCommand())
                    {
                        string sFinish_Date = "";
                        //if (m_QAMaster.Status == "Closed")
                        if (m_QAMaster.Status == "273")
                        {
                            sFinish_Date = "Finish_Date=(Convert(varchar(10),Getdate(),111)) ,Finish_Time=(Convert(varchar(8),Getdate(),108)),";
                        }
                        //else
                        //{ sFinish_Date = "Finish_Date='' ,"; }

                        //string[] sArrCase_Owner;
                        //ReturnParameter(out sArrCase_Owner, m_QAMaster.System);
                        //string sCase_Owner = sArrCase_Owner[0];

                        if (m_QAMaster.DocID == "")
                        { m_QAMaster.DocID = DateTime.Now.ToString("yyyyMMddHHmmssfff"); }

                        m_SQL = "Update [dbo].[QAMaster]  set " +
                            "Priority =@Priority, " +
                            "System1 =@System1 ," +
                            "System3 =@System3 ," +
                            "DocID =@DocID," +
                            "QASubject =@QASubject," +
                            "QADescription =@QADescription," +
                            "QAReply =@QAReply," +
                            "KeyWords =@KeyWords," +
                            "Status =@Status," +
                            sFinish_Date +
                            //"Case_Owner =@Case_Owner," +
                            "Issue_Type =@Issue_Type," +
                            "ModifyUser=@ModifyUser , " +
                            "ModifyDate=(Convert(varchar(10),Getdate(),111)), ModifyTime=(Convert(varchar(8),Getdate(),108)) ";

                        m_cmd.Parameters.Add(new SqlParameter("Priority", m_QAMaster.Priority));
                        m_cmd.Parameters.Add(new SqlParameter("System1", m_QAMaster.System1));
                        m_cmd.Parameters.Add(new SqlParameter("System3", m_QAMaster.System3));
                        m_cmd.Parameters.Add(new SqlParameter("DocID", m_QAMaster.DocID));
                        m_cmd.Parameters.Add(new SqlParameter("QASubject", m_QAMaster.QASubject));
                        m_cmd.Parameters.Add(new SqlParameter("QADescription", m_QAMaster.QADescription));
                        m_cmd.Parameters.Add(new SqlParameter("QAReply", m_QAMaster.QAReply));
                        m_cmd.Parameters.Add(new SqlParameter("KeyWords", m_QAMaster.KeyWords));
                        m_cmd.Parameters.Add(new SqlParameter("Status", m_QAMaster.Status));
                        //m_cmd.Parameters.Add(new SqlParameter("Case_Owner", sCase_Owner));
                        m_cmd.Parameters.Add(new SqlParameter("Issue_Type", m_QAMaster.Issue_Type));
                        m_cmd.Parameters.Add(new SqlParameter("ModifyUser", p_EmployeeEntity.WorkID));

                        m_SQL += " where [QAID]=@QAID ";
                        m_cmd.Parameters.Add(new SqlParameter("QAID", m_QAMaster.QAID));
                        m_cmd.CommandText = m_SQL;
                        m_TransResult = g_DC.ExecuteNonQuery(p_EmployeeEntity, m_cmd);

                        if (m_TransResult.isSuccess)
                        {
                            SysEntity.SmtpMail mail = new SysEntity.SmtpMail();
                          
                          
                        }

                    }
                }
                catch (Exception ex)
                {
                    m_TransResult.LogMessage = ex.Message;
                    m_TransResult.isSuccess = false;
                }
            }
            return m_TransResult;
        }

        public SysEntity.TransResult GetSiteInfo(SysEntity.Employee p_EmployeeEntity, string SiteID)
        {
            SysEntity.TransResult m_TransResult = new SysEntity.TransResult();
            try
            {
                using (SqlCommand m_cmd = new SqlCommand("SELECT * FROM [dbo].[Site] Where SiteID=@SiteID "))
                {
                    var m_Params = new
                    {
                        SiteID = SiteID
                    };

                    m_TransResult = g_DC.GetData<SysEntity.Site>(p_EmployeeEntity, m_cmd, m_Params);
                    if (m_TransResult.isSuccess)
                    {
                        List<SysEntity.Site> m_lisSite = (List<SysEntity.Site>)m_TransResult.ResultEntity;

                        if (m_lisSite.Count != 0)
                        {
                            foreach (SysEntity.Site Empl in m_lisSite)
                            {
                                m_TransResult.ResultEntity = Empl;
                            }
                        }
                        else
                        {
                            m_TransResult.isSuccess = false;
                            m_TransResult.LogMessage = "Site ID:" + SiteID + ",data is empty";
                        }

                    }
                }
            }
            catch (Exception ex)
            {
                m_TransResult.LogMessage = ex.Message;
                m_TransResult.isSuccess = false;
            }
            return m_TransResult;
        }

        public SysEntity.TransResult DelQAMaster(SysEntity.Employee p_EmployeeEntity, Dictionary<string, Object> p_dicQAMaster)
        {
            SysEntity.TransResult m_TransResult = new SysEntity.TransResult();
            Dictionary<string, Object> m_ArrayList = new Dictionary<string, Object>();
            SysEntity.QAMaster m_QAMaster = g_FH.ConvertEntity<SysEntity.QAMaster>(p_dicQAMaster, ref m_ArrayList);

            try
            {
                string m_SQL = "Delete [dbo].[QAMaster] Where [QAID]=@QAID";
                using (SqlCommand m_cmd = new SqlCommand(m_SQL))
                {
                    m_cmd.Parameters.Add(new SqlParameter("QAID", m_QAMaster.QAID));
                    m_TransResult = g_DC.ExecuteNonQuery(p_EmployeeEntity, m_cmd);
                }
            }
            catch (Exception ex)
            {
                m_TransResult.LogMessage = ex.Message;
                m_TransResult.isSuccess = false;
            }
            return m_TransResult;
        }

        public SysEntity.TransResult GetDocID(SysEntity.Employee p_EmployeeEntity)
        {
            SysEntity.TransResult m_TransResult = new SysEntity.TransResult();
            try
            {
                m_TransResult.isSuccess = true;
                m_TransResult.ResultEntity = DateTime.Now.ToString("yyyyMMddHHmmssfff");
            }
            catch (Exception ex)
            {
                m_TransResult.LogMessage = ex.Message;
                m_TransResult.isSuccess = false;
            }

            return m_TransResult;
        }

        public SysEntity.TransResult GetUserDepartment(SysEntity.Employee p_EmployeeEntity)
        {
            SysEntity.TransResult m_TransResult = new SysEntity.TransResult();
            try
            {
                string m_SQL = " select ORGEH,SHORT1 from [dbo].[EmployeeInfo] E Left Join (SELECT [OBJID],[SHORT1] FROM [dbo].[DepartmentInfo] ) D on E.ORGEH=D.OBJID where E.PERNR=@PERNR  ";
                using (SqlCommand m_cmd = new SqlCommand(m_SQL))
                {
                    m_cmd.Parameters.Add(new SqlParameter("PERNR", p_EmployeeEntity.WorkID));
                    m_TransResult = g_DC.GetDataTable(p_EmployeeEntity, m_cmd);
                }
                if (m_TransResult.isSuccess)
                {
                    foreach (DataRow dr in ((DataTable)m_TransResult.ResultEntity).Rows)
                    {
                        m_TransResult.ResultEntity = dr["SHORT1"].ToString();
                    }
                }
            }
            catch (Exception ex)
            {
                m_TransResult.LogMessage = ex.Message;
                m_TransResult.isSuccess = false;
            }

            return m_TransResult;
        }

        public SysEntity.TransResult GetUserName(SysEntity.Employee p_EmployeeEntity)
        {
            SysEntity.TransResult m_TransResult = new SysEntity.TransResult();
            try
            {
                Dictionary<string, Object> m_disReturn = new Dictionary<string, object>();
                string m_SQL = " SELECT @PERNR as User_Contact,  dbo.fun_GetEmployeeNameExt(@PERNR) EmployeeName ";
                using (SqlCommand m_cmd = new SqlCommand(m_SQL))
                {
                    m_cmd.Parameters.Add(new SqlParameter("PERNR", p_EmployeeEntity.WorkID));
                    m_TransResult = g_DC.GetDataTable(p_EmployeeEntity, m_cmd);
                }
                if (m_TransResult.isSuccess)
                {
                    DataTable m_dt = ((DataTable)m_TransResult.ResultEntity);
                    foreach (DataRow dr in m_dt.Rows)
                    {
                        foreach (DataColumn dc in m_dt.Columns)
                        {
                            m_disReturn[dc.ColumnName] = dr[dc.ColumnName].ToString();
                        }
                    }
                    m_TransResult.ResultEntity = m_disReturn;
                }
            }
            catch (Exception ex)
            {
                m_TransResult.LogMessage = ex.Message;
                m_TransResult.isSuccess = false;
            }

            return m_TransResult;
        }

        public SysEntity.TransResult FormTransferQA(SysEntity.Employee p_EmployeeEntity, Dictionary<string, Object> p_DicFormTransfer)
        {
            SysEntity.TransResult m_TransResult = new SysEntity.TransResult();
            try
            {
                //p_DicFormTransfer["WorkID"].ToString(); 轉送給誰
                //p_DicFormTransfer["TransferRemark"].ToString(); 轉送備註
                //p_DicFormTransfer["FormKeys"].ToString(); 轉送Key
                List<Business_Logic.Entity.SysEntity.CommandEntity> m_MultiCommandEntitys = new List<SysEntity.CommandEntity>();

                m_TransResult = g_FH.CheckFormKeys(p_DicFormTransfer);
                if (m_TransResult.isSuccess)
                {
                    Dictionary<string, Object> m_Params = (Dictionary<string, Object>)m_TransResult.ResultEntity;
                    List<object> m_lisParams = new List<object>();

                    Business_Logic.Entity.SysEntity.CommandEntity m_MultiCommandEntity = new SysEntity.CommandEntity();
                    object m_Param = new
                    {
                        QAID = m_Params["QAID"].ToString(),
                        Case_Owner = p_DicFormTransfer["WorkID"].ToString()
                    };
                    m_MultiCommandEntity.SQLCommand = "update [dbo].[QAMaster] set Case_Owner=@Case_Owner  Where [QAID]=@QAID";
                    m_MultiCommandEntity.Parameters = m_Param;
                    m_MultiCommandEntitys.Add(m_MultiCommandEntity);

                    p_DicFormTransfer["TransferKey"] = m_Params["QAID"];
                    m_TransResult = CommitFormTransfer(p_EmployeeEntity, p_DicFormTransfer, m_MultiCommandEntitys);

                   

                }
            }
            catch (Exception ex)
            {
                m_TransResult.LogMessage = ex.Message;
                m_TransResult.isSuccess = false;
            }
            return m_TransResult;
        }

        public SysEntity.TransResult CancelQA(SysEntity.Employee p_EmployeeEntity, Dictionary<string, Object> p_DicFormTransfer)
        {
            SysEntity.TransResult m_TransResult = new SysEntity.TransResult();
            try
            {
                string m_QADescription = p_DicFormTransfer["CancelRemark"].ToString();// 註銷備註
                m_TransResult = g_FH.CheckFormKeys(p_DicFormTransfer);
                if (m_TransResult.isSuccess)
                {
                    Dictionary<string, Object> m_Params = (Dictionary<string, Object>)m_TransResult.ResultEntity;
                    string m_SQL = "update [dbo].[QAMaster] set Status='274' ,QADescription=QADescription+char(10)+char(13)+@QADescription  Where [QAID]=@QAID";
                    using (SqlCommand m_cmd = new SqlCommand(m_SQL))
                    {
                        m_cmd.Parameters.Add(new SqlParameter("QAID", m_Params["QAID"].ToString()));
                        m_cmd.Parameters.Add(new SqlParameter("QADescription", m_QADescription));

                        m_TransResult = g_DC.ExecuteNonQuery(p_EmployeeEntity, m_cmd);
                    }

                    if (m_TransResult.isSuccess)
                    {
                        //SysEntity.SmtpMail mail = new SysEntity.SmtpMail();
                        //mail.MailTo = getEmail(p_EmployeeEntity, p_DicFormTransfer["WorkID"].ToString());
                        ////mail.MailTo = "Danny_Chu@adata.com";
                        //string sSubject = "";

                        ////sSubject = "轉送通知：Issue Report：" + m_Params["QAID"].ToString() + "，申請人：" + p_DicFormTransfer["WorkID"].ToString();
                        //string[] sArrMsg;
                        //GetQAMsg(out sArrMsg, m_Params["QAID"].ToString());
                        //sSubject = "註銷通知：Issue Report：" + m_Params["QAID"].ToString() + "，申請人：" + sArrMsg[0].ToString();

                        //mail.Subject = sSubject;
                        //Business_Logic.Entity.SysEntity.HyperlinkEntity p_HyperlinkEntity = new SysEntity.HyperlinkEntity();
                        //p_HyperlinkEntity.SiteFormID = "1000SYS9001";
                        //Dictionary<string, string> m_KeyValues = new Dictionary<string, string>();
                        //m_KeyValues["QAID"] = m_Params["QAID"].ToString();
                        //p_HyperlinkEntity.KeyValues = m_KeyValues;
                        //string m_URL = g_FH.GetMaintainHyperlink(p_HyperlinkEntity);
                        //mail.Body = sSubject + "<br>請點選下方連結..<a href='https://" + m_URL + "'>Issue Report</a>(含行動簽核連結)<br>";

                        //try
                        //{ g_FH.SendMail(mail); }
                        //catch (Exception ex)
                        //{ m_TransResult.LogMessage = ex.Message; }
                    }

                }
            }
            catch (Exception ex)
            {
                m_TransResult.LogMessage = ex.Message;
                m_TransResult.isSuccess = false;
            }
            return m_TransResult;
        }

        public void GetQAMsg(out string[] arr, string QAID)
        {
            arr = new string[2];

            string m_SQL = "SELECT dbo.fun_GetEmployeeNameExt(User_Contact) as User_ContactALL," +
                "dbo.fun_GetEmployeeNameExt(Case_Owner) as Case_OwnerALL " +
                " FROM [dbo].[QAMaster] WHERE [QAID]=@QAID   ";

            using (SqlCommand m_cmd = new SqlCommand(m_SQL))
            {
                m_cmd.Parameters.Add(new SqlParameter("QAID", QAID));
                DataView dvResult = g_DC.GetDataTable(m_cmd).DefaultView;
                if (dvResult != null)
                {
                    if (dvResult.Count > 0)
                    {
                        for (int i = 0; i < 2; i++)
                        { arr[i] = dvResult.Table.Rows[0][i].ToString(); }
                    }
                }
            }
        }

        public SysEntity.TransResult getFormTransferLog(SysEntity.Employee p_EmployeeEntity, Dictionary<string, Object> p_dicQAMaster)
        {
            SysEntity.TransResult m_TransResult = new SysEntity.TransResult();
            Dictionary<string, Object> m_ArrayList = new Dictionary<string, Object>();
            SysEntity.QAMaster m_QAMaster = g_FH.ConvertEntity<SysEntity.QAMaster>(p_dicQAMaster, ref m_ArrayList);

            try
            {
                string m_SQL = " SELECT   FormTransferID, dbo.fun_GetEmployeeNameExt(TransferUserFrom) as TransferUserFrom ,dbo.fun_GetEmployeeNameExt(TransferUserTo) as TransferUserTo  ,TransferRemark,CONVERT (datetime,substring(FormTransferID,1,8 )+' ' +substring(FormTransferID,9,2 )+':' +substring(FormTransferID,11,2 )+':' +substring(FormTransferID,13,2 ), 120) as CreateDateTime" +
                    " FROM              FormTransferLog  where TransferKey=@TransferKey  ";
                using (SqlCommand m_cmd = new SqlCommand(m_SQL))
                {
                    m_cmd.Parameters.Add(new SqlParameter("TransferKey", m_QAMaster.QAID));
                    //m_cmd.Parameters.Add(new SqlParameter("TransferKey", "QA170900001"));
                    m_cmd.CommandText += " order by [FormTransferID] Desc ";
                    m_TransResult = g_DC.GetDataTable(p_EmployeeEntity, m_cmd);
                }
            }
            catch (Exception ex)
            {
                m_TransResult.LogMessage = ex.Message;
                m_TransResult.isSuccess = false;
            }
            return m_TransResult;
        }

        public SysEntity.TransResult GetQA_System_A(SysEntity.Employee p_EmployeeEntity, Dictionary<string, Object> p_dicQAMaster)
        {
            SysEntity.TransResult m_TransResult = new SysEntity.TransResult();
            Dictionary<string, Object> m_objBetween = new Dictionary<string, Object>();
            SysEntity.QAMaster m_QAMaster = g_FH.ConvertEntity<SysEntity.QAMaster>(p_dicQAMaster, ref m_objBetween);
            try
            {
                string sCurrentLanguage = "TW";
                try
                {
                    if (p_EmployeeEntity.CurrentLanguage.ToString() == "")
                    { sCurrentLanguage = "TW"; }
                    else
                    { sCurrentLanguage = p_EmployeeEntity.CurrentLanguage.ToString(); }
                }
                catch
                { sCurrentLanguage = "TW"; }

                string m_SQL = "SELECT sValue AS Value1, sText" + sCurrentLanguage + " AS Text1,OrderIndex  FROM dbo.MultiParameter WHERE (ParamID = 'QA_SystemH')  ";
                using (SqlCommand m_cmd = g_DC.GetQueryCommand(p_EmployeeEntity, m_SQL, m_QAMaster, m_objBetween))
                {
                    m_cmd.CommandText += " order by [OrderIndex]  ";
                    m_TransResult = g_DC.GetDataTable(p_EmployeeEntity, m_cmd);
                }
            }
            catch (Exception ex)
            {
                m_TransResult.LogMessage = ex.Message;
                m_TransResult.isSuccess = false;
            }
            return m_TransResult;
        }
        public SysEntity.TransResult GetQA_System_B(SysEntity.Employee p_EmployeeEntity, Dictionary<string, Object> p_dicQAMaster)
        {
            SysEntity.TransResult m_TransResult = new SysEntity.TransResult();
            Dictionary<string, Object> m_objBetween = new Dictionary<string, Object>();
            SysEntity.QAMaster m_QAMaster = g_FH.ConvertEntity<SysEntity.QAMaster>(p_dicQAMaster, ref m_objBetween);
            try
            {

                string sCurrentLanguage = "TW";
                string sSystem1 = "";
                try
                {
                    if (p_EmployeeEntity.CurrentLanguage.ToString() == "")
                    { sCurrentLanguage = "TW"; }
                    else
                    { sCurrentLanguage = p_EmployeeEntity.CurrentLanguage.ToString(); }

                    if (m_QAMaster.System1.ToString() == "")
                    { sSystem1 = ""; }
                    else
                    { sSystem1 = m_QAMaster.System1.ToString(); }


                }
                catch
                {
                    sCurrentLanguage = "TW";
                    sSystem1 = "";
                }

                string m_SQL = "SELECT dbo.MultiParameter.sFatherV as System1, dbo.MultiParameter.sValue as Value2,dbo.MultiParameter.sText" + sCurrentLanguage + " as Value3, dbo.MultiParameter.OrderIndex " +
                "FROM dbo.MultiParameter LEFT OUTER JOIN dbo.MultiParameter AS MultiParameter_1 ON dbo.MultiParameter.sFatherV = MultiParameter_1.QASystemSN " +
                "WHERE (dbo.MultiParameter.ParamID = 'QA_SystemD')  and dbo.MultiParameter.sFatherV='" + sSystem1 + "' ";

                using (SqlCommand m_cmd = g_DC.GetQueryCommand(p_EmployeeEntity, m_SQL, m_QAMaster, m_objBetween))
                {
                    m_cmd.CommandText += " order by [OrderIndex]  ";
                    m_TransResult = g_DC.GetDataTable(p_EmployeeEntity, m_cmd);
                }
            }
            catch (Exception ex)
            {
                m_TransResult.LogMessage = ex.Message;
                m_TransResult.isSuccess = false;
            }
            return m_TransResult;
        }
        public SysEntity.TransResult GetQA_System_C(SysEntity.Employee p_EmployeeEntity, Dictionary<string, Object> p_dicQAMaster)
        {
            SysEntity.TransResult m_TransResult = new SysEntity.TransResult();
            Dictionary<string, Object> m_objBetween = new Dictionary<string, Object>();
            SysEntity.QAMaster m_QAMaster = g_FH.ConvertEntity<SysEntity.QAMaster>(p_dicQAMaster, ref m_objBetween);
            try
            {
                string sCurrentLanguage = "TW";
                try
                {
                    if (p_EmployeeEntity.CurrentLanguage.ToString() == "")
                    { sCurrentLanguage = "TW"; }
                    else
                    { sCurrentLanguage = p_EmployeeEntity.CurrentLanguage.ToString(); }
                }
                catch
                { sCurrentLanguage = "TW"; }

                string m_SQL = "SELECT dbo.MultiParameter.sValue as Value2,dbo.MultiParameter.sText" + sCurrentLanguage + " as Value3, dbo.MultiParameter.OrderIndex " +
                "FROM dbo.MultiParameter LEFT OUTER JOIN dbo.MultiParameter AS MultiParameter_1 ON dbo.MultiParameter.sFatherV = MultiParameter_1.QASystemSN " +
                "WHERE (dbo.MultiParameter.ParamID = 'QA_SystemD')  ";

                using (SqlCommand m_cmd = g_DC.GetQueryCommand(p_EmployeeEntity, m_SQL, m_QAMaster, m_objBetween))
                {
                    m_cmd.CommandText += " order by [OrderIndex]  ";
                    m_TransResult = g_DC.GetDataTable(p_EmployeeEntity, m_cmd);
                }
            }
            catch (Exception ex)
            {
                m_TransResult.LogMessage = ex.Message;
                m_TransResult.isSuccess = false;
            }
            return m_TransResult;
        }
        private void ReturnQASystem(out string[] arr, string QASystemSN)
        {
            arr = new string[7];
            string m_SQL = "SELECT ParamID,sFatherV,sTextTW,sTextUS,sValue,OrderIndex,Owner1 FROM [dbo].[MultiParameter] WHERE QASystemSN=@QASystemSN   ";
            using (SqlCommand m_cmd = new SqlCommand(m_SQL))
            {
                m_cmd.Parameters.Add(new SqlParameter("QASystemSN", QASystemSN));
                DataView dvResult = g_DC.GetDataTable(m_cmd).DefaultView;
                if (dvResult != null)
                {
                    if (dvResult.Count > 0)
                    {
                        for (int i = 0; i < 7; i++)
                        { arr[i] = dvResult.Table.Rows[0][i].ToString(); }
                    }
                }
            }
        }

        public SysEntity.TransResult GetQA_Site_A(SysEntity.Employee p_EmployeeEntity, Dictionary<string, Object> p_dicParameter)
        {
            SysEntity.TransResult m_TransResult = new SysEntity.TransResult();
            Dictionary<string, Object> m_objBetween = new Dictionary<string, Object>();
            SysEntity.Parameter m_Parameter = g_FH.ConvertEntity<SysEntity.Parameter>(p_dicParameter, ref m_objBetween);
            try
            {
                string sCurrentLanguage = "TW";
                try
                {
                    if (p_EmployeeEntity.CurrentLanguage.ToString() == "")
                    { sCurrentLanguage = "TW"; }
                    else
                    { sCurrentLanguage = p_EmployeeEntity.CurrentLanguage.ToString(); }
                }
                catch
                { sCurrentLanguage = "TW"; }

                string sMyWord = "Description";
                if (sCurrentLanguage != "TW")
                {
                    sMyWord = "DescriptionEnglish";
                }

                string m_SQL = "SELECT '' as Value , '' as Text  FROM [dbo].[Site] union SELECT Description as Value ," + sMyWord + " as Text   FROM [dbo].[Site]   ";
                using (SqlCommand m_cmd = g_DC.GetQueryCommand(p_EmployeeEntity, m_SQL, m_Parameter, m_objBetween))
                {
                    //m_cmd.CommandText += " order by [SiteID]  ";
                    m_TransResult = g_DC.GetDataTable(p_EmployeeEntity, m_cmd);
                }
            }
            catch (Exception ex)
            {
                m_TransResult.LogMessage = ex.Message;
                m_TransResult.isSuccess = false;
            }
            return m_TransResult;
        }

        public SysEntity.TransResult GetQA_Issue_Type_0A(SysEntity.Employee p_EmployeeEntity, Dictionary<string, Object> p_dicParameter)
        {
            SysEntity.TransResult m_TransResult = new SysEntity.TransResult();
            Dictionary<string, Object> m_objBetween = new Dictionary<string, Object>();
            SysEntity.Parameter m_Parameter = g_FH.ConvertEntity<SysEntity.Parameter>(p_dicParameter, ref m_objBetween);
            try
            {
                string sCurrentLanguage = "TW";
                try
                {
                    if (p_EmployeeEntity.CurrentLanguage.ToString() == "")
                    { sCurrentLanguage = "TW"; }
                    else
                    { sCurrentLanguage = p_EmployeeEntity.CurrentLanguage.ToString(); }
                }
                catch
                { sCurrentLanguage = "TW"; }

                //string m_SQL = "SELECT * FROM [dbo].[Parameter]  Where  ParamID='QA_Issue_Type" + p_EmployeeEntity.CurrentLanguage + "' ";
                //string m_SQL = "SELECT QASystemSN as Value ,sText" + p_EmployeeEntity.CurrentLanguage + " as Text,OrderIndex  FROM [dbo].[QASystem]  Where ParamID = 'QA_Issue_Type'  ";
                string m_SQL = MkSQL(sCurrentLanguage, "QA_Issue_Type", "0");
                using (SqlCommand m_cmd = g_DC.GetQueryCommand(p_EmployeeEntity, m_SQL, m_Parameter, m_objBetween))
                {
                    m_cmd.CommandText += " order by [OrderIndex]  ";
                    m_TransResult = g_DC.GetDataTable(p_EmployeeEntity, m_cmd);
                }
            }
            catch (Exception ex)
            {
                m_TransResult.LogMessage = ex.Message;
                m_TransResult.isSuccess = false;
            }
            return m_TransResult;
        }
        public SysEntity.TransResult GetQA_Priority_0A(SysEntity.Employee p_EmployeeEntity, Dictionary<string, Object> p_dicParameter)
        {
            SysEntity.TransResult m_TransResult = new SysEntity.TransResult();
            Dictionary<string, Object> m_objBetween = new Dictionary<string, Object>();
            SysEntity.Parameter m_Parameter = g_FH.ConvertEntity<SysEntity.Parameter>(p_dicParameter, ref m_objBetween);
            try
            {
                string sCurrentLanguage = "TW";
                try
                {
                    if (p_EmployeeEntity.CurrentLanguage.ToString() == "")
                    { sCurrentLanguage = "TW"; }
                    else
                    { sCurrentLanguage = p_EmployeeEntity.CurrentLanguage.ToString(); }
                }
                catch
                { sCurrentLanguage = "TW"; }

                string m_SQL = MkSQL(sCurrentLanguage, "QA_Priority", "0");
                using (SqlCommand m_cmd = g_DC.GetQueryCommand(p_EmployeeEntity, m_SQL, m_Parameter, m_objBetween))
                {
                    m_cmd.CommandText += " order by [OrderIndex]  ";
                    m_TransResult = g_DC.GetDataTable(p_EmployeeEntity, m_cmd);
                }
            }
            catch (Exception ex)
            {
                m_TransResult.LogMessage = ex.Message;
                m_TransResult.isSuccess = false;
            }
            return m_TransResult;
        }
        public SysEntity.TransResult GetQA_Status_0A(SysEntity.Employee p_EmployeeEntity, Dictionary<string, Object> p_dicParameter)
        {
            SysEntity.TransResult m_TransResult = new SysEntity.TransResult();
            Dictionary<string, Object> m_objBetween = new Dictionary<string, Object>();
            SysEntity.Parameter m_Parameter = g_FH.ConvertEntity<SysEntity.Parameter>(p_dicParameter, ref m_objBetween);
            try
            {
                string sCurrentLanguage = "TW";
                try
                {
                    if (p_EmployeeEntity.CurrentLanguage.ToString() == "")
                    { sCurrentLanguage = "TW"; }
                    else
                    { sCurrentLanguage = p_EmployeeEntity.CurrentLanguage.ToString(); }
                }
                catch
                { sCurrentLanguage = "TW"; }

                string m_SQL = MkSQL(sCurrentLanguage, "QA_Status", "0");
                using (SqlCommand m_cmd = g_DC.GetQueryCommand(p_EmployeeEntity, m_SQL, m_Parameter, m_objBetween))
                {
                    m_cmd.CommandText += " order by [OrderIndex]  ";
                    m_TransResult = g_DC.GetDataTable(p_EmployeeEntity, m_cmd);
                }
            }
            catch (Exception ex)
            {
                m_TransResult.LogMessage = ex.Message;
                m_TransResult.isSuccess = false;
            }
            return m_TransResult;
        }
        public SysEntity.TransResult GetQA_Issue_TypeA(SysEntity.Employee p_EmployeeEntity, Dictionary<string, Object> p_dicParameter)
        {
            SysEntity.TransResult m_TransResult = new SysEntity.TransResult();
            Dictionary<string, Object> m_objBetween = new Dictionary<string, Object>();
            SysEntity.Parameter m_Parameter = g_FH.ConvertEntity<SysEntity.Parameter>(p_dicParameter, ref m_objBetween);
            try
            {
                string sCurrentLanguage = "TW";
                try
                {
                    if (p_EmployeeEntity.CurrentLanguage.ToString() == "")
                    { sCurrentLanguage = "TW"; }
                    else
                    { sCurrentLanguage = p_EmployeeEntity.CurrentLanguage.ToString(); }
                }
                catch
                { sCurrentLanguage = "TW"; }

                string m_SQL = MkSQL(sCurrentLanguage, "QA_Issue_Type", "");
                using (SqlCommand m_cmd = g_DC.GetQueryCommand(p_EmployeeEntity, m_SQL, m_Parameter, m_objBetween))
                {
                    m_cmd.CommandText += " order by [OrderIndex]  ";
                    m_TransResult = g_DC.GetDataTable(p_EmployeeEntity, m_cmd);
                }
            }
            catch (Exception ex)
            {
                m_TransResult.LogMessage = ex.Message;
                m_TransResult.isSuccess = false;
            }
            return m_TransResult;
        }
        public SysEntity.TransResult GetQA_PriorityA(SysEntity.Employee p_EmployeeEntity, Dictionary<string, Object> p_dicParameter)
        {
            SysEntity.TransResult m_TransResult = new SysEntity.TransResult();
            Dictionary<string, Object> m_objBetween = new Dictionary<string, Object>();
            SysEntity.Parameter m_Parameter = g_FH.ConvertEntity<SysEntity.Parameter>(p_dicParameter, ref m_objBetween);
            try
            {
                string sCurrentLanguage = "TW";
                try
                {
                    if (p_EmployeeEntity.CurrentLanguage.ToString() == "")
                    { sCurrentLanguage = "TW"; }
                    else
                    { sCurrentLanguage = p_EmployeeEntity.CurrentLanguage.ToString(); }
                }
                catch
                { sCurrentLanguage = "TW"; }

                string m_SQL = MkSQL(sCurrentLanguage, "QA_Priority", "");
                using (SqlCommand m_cmd = g_DC.GetQueryCommand(p_EmployeeEntity, m_SQL, m_Parameter, m_objBetween))
                {
                    m_cmd.CommandText += " order by [OrderIndex]  ";
                    m_TransResult = g_DC.GetDataTable(p_EmployeeEntity, m_cmd);
                }
            }
            catch (Exception ex)
            {
                m_TransResult.LogMessage = ex.Message;
                m_TransResult.isSuccess = false;
            }
            return m_TransResult;
        }
        public SysEntity.TransResult GetQA_StatusA(SysEntity.Employee p_EmployeeEntity, Dictionary<string, Object> p_dicParameter)
        {
            SysEntity.TransResult m_TransResult = new SysEntity.TransResult();
            Dictionary<string, Object> m_objBetween = new Dictionary<string, Object>();
            SysEntity.Parameter m_Parameter = g_FH.ConvertEntity<SysEntity.Parameter>(p_dicParameter, ref m_objBetween);
            try
            {
                string sCurrentLanguage = "TW";
                try
                {
                    if (p_EmployeeEntity.CurrentLanguage.ToString() == "")
                    { sCurrentLanguage = "TW"; }
                    else
                    { sCurrentLanguage = p_EmployeeEntity.CurrentLanguage.ToString(); }
                }
                catch
                { sCurrentLanguage = "TW"; }

                string m_SQL = MkSQL(sCurrentLanguage, "QA_Status", "");
                using (SqlCommand m_cmd = g_DC.GetQueryCommand(p_EmployeeEntity, m_SQL, m_Parameter, m_objBetween))
                {
                    m_cmd.CommandText += " order by [OrderIndex]  ";
                    m_TransResult = g_DC.GetDataTable(p_EmployeeEntity, m_cmd);
                }
            }
            catch (Exception ex)
            {
                m_TransResult.LogMessage = ex.Message;
                m_TransResult.isSuccess = false;
            }
            return m_TransResult;
        }
        private string MkSQL(string MyLang, string KeyWord, string s01YN)
        {
            if (s01YN != "0")
            { s01YN = "AND OrderIndex > '01'  "; }
            else
            { s01YN = ""; }

            string sOkSql = "";
            sOkSql = "SELECT sValue as Value ,sText" + MyLang + " as Text,OrderIndex  FROM MultiParameter Where ParamID = '" + KeyWord + "'  " + s01YN;
            return sOkSql;
        }
        
        public Object CheckTransResult<T>(SysEntity.TransResult p_TransResult) where T : new()
        {
            SysEntity.TransResult m_TransResult = new SysEntity.TransResult();
            Object m_Result = new Object();

            try
            {
                if (p_TransResult != null)
                {
                    if (p_TransResult.isSuccess)
                    {
                        var m_STypet = (T)Convert.ChangeType(p_TransResult.ResultEntity, typeof(T));
                        m_Result = m_STypet;
                    }
                    else
                    {
                        var m_STypet = p_TransResult.LogMessage;
                        m_Result = m_STypet;

                    }
                }

            }
            catch (Exception ex)
            {
                throw ex;
            }
            return m_Result;
        }

        //掛號程式 For QAID
        private string RegisterQAID()
        {
            string sNewQAID = "";
            string m_SQL = "SELECT (CASE When MAX(QAID) IS NULL Then 'QA' + SUBSTRING(CONVERT(char(6), GETDATE(), 112), 3, 4)+'00000' else MAX(QAID) End ) AS sNewQAID FROM [dbo].[QAMaster] WHERE (LEFT(QAID, 6) = 'QA' + SUBSTRING(CONVERT(char(6), GETDATE(), 112), 3, 4)) ";
            using (SqlCommand m_cmd = new SqlCommand(m_SQL))
            {
                DataView dvResult = g_DC.GetDataTable(m_cmd).DefaultView;
                if (dvResult != null)
                {
                    if (dvResult.Count > 0)
                    { sNewQAID = "QA" + (Convert.ToInt64(dvResult.Table.Rows[0]["sNewQAID"].ToString().Substring(2, 9)) + 1).ToString(); }
                }
                if (sNewQAID == "")
                { sNewQAID = "QA" + DateTime.Now.ToString("yyMM") + "00001"; }
            }

            string m_SQL2 = "INSERT INTO dbo.[QAMaster] ( QAID ) Values ( @QAID ) ";
            using (SqlCommand m_cmd = new SqlCommand(m_SQL2))
            {
                m_cmd.Parameters.Add(new SqlParameter("QAID", sNewQAID));
                SysEntity.Employee p_EmployeeEntity = new SysEntity.Employee();
                g_DC.ExecuteNonQuery(p_EmployeeEntity, m_cmd);
            }
            return sNewQAID;
        }

        #endregion





        #region 取得Menu





        public SysEntity.TransResult GetMainMenuInfo(SysEntity.Employee p_EmployeeEntity, string MenuType, string SiteID ,string MenuID)
        {
            SysEntity.TransResult m_TransResult = new SysEntity.TransResult();
            try
            {
                string m_SQL = " select M.*,I.ModuleName MenuDesc,F.Description FormDesc from [MainMenu] M   left join ModuleInfo I on M.ModuleKey=I.ModuleKey " +
                    " left join FormInfo F  on M.FormID=F.FormID " +
                    "  where MenuType =@MenuType and SiteID =@SiteID  ";
                object m_Params = new object();
                if (MenuID != "0")
                {
                    m_SQL += " and MenuParentID=@MenuParentID";

                     m_Params = new
                    {
                        SiteID,
                        MenuType,
                        MenuParentID = MenuID

                    };
                }
                else
                {
                     m_Params = new
                    {
                        SiteID,
                        MenuType,
                    };
                  
                }
                m_SQL += " order by MenuIndex asc ";


                using (SqlCommand m_cmd = new SqlCommand(m_SQL))
                {
                    m_TransResult = g_DC.GetData<SysEntity.MainMenu>(p_EmployeeEntity, m_cmd, m_Params);

                    if (m_TransResult.isSuccess)
                    {
                        List<SysEntity.MainMenu> m_lisMainMenu = (List<SysEntity.MainMenu>)m_TransResult.ResultEntity;
                        m_TransResult.ResultEntity = m_lisMainMenu;
                    }
                }
            }
            catch (Exception ex)
            {
                m_TransResult.LogMessage = ex.Message;
                m_TransResult.isSuccess = false;
            }
            return m_TransResult;
        }

        public SysEntity.TransResult GetAuthMenuInfo(SysEntity.Employee p_EmployeeEntity,  string MenuID)
        {
            SysEntity.TransResult m_TransResult = new SysEntity.TransResult();
            try
            {
                string m_SQL = " SELECT  [MenuID] ,A.[GroupID],G.GroupDesc FROM [dbo].[AuthMenu] A Left Join View_GroupID G on A.GroupID=G.GroupID   where MenuID =@MenuID ";
                  var  m_Params = new
                    {
                      MenuID,
                    };
              
                using (SqlCommand m_cmd = new SqlCommand(m_SQL))
                {
                    m_TransResult = g_DC.GetData<SysEntity.AuthMenu>(p_EmployeeEntity, m_cmd, m_Params);
                    if (m_TransResult.isSuccess)
                    {
                        List<SysEntity.AuthMenu> m_lisAuthMenu = (List<SysEntity.AuthMenu>)m_TransResult.ResultEntity;
                        m_TransResult.ResultEntity = m_lisAuthMenu;
                    }
                }
            }
            catch (Exception ex)
            {
                m_TransResult.LogMessage = ex.Message;
                m_TransResult.isSuccess = false;
            }
            return m_TransResult;
        }


        

        public SysEntity.TransResult GetModuleKey(SysEntity.Employee p_EmployeeEntity, string SiteID, string ActType, string MenuType)
        {
            SysEntity.TransResult m_TransResult = new SysEntity.TransResult();
            try
            {
                string m_Condition = " in ";
                if (ActType == "add" && MenuType=="M")
                {
                    m_Condition = " not in";
                }
                using (SqlCommand m_cmd = new SqlCommand(" select [ModuleKey],[ModuleName] from [ModuleInfo] where [ModuleKey] " + m_Condition + " (SELECT distinct [ModuleKey] FROM [dbo].[MainMenu] where [SiteID]=@SiteID) "))
                {
                    var m_Params = new
                    {
                        SiteID = SiteID
                    };

                    m_TransResult = g_DC.GetData<SysEntity.ModuleInfo>(p_EmployeeEntity, m_cmd, m_Params);
                    if (m_TransResult.isSuccess)
                    {
                        List<SysEntity.ModuleInfo> m_lisModuleInfo = (List<SysEntity.ModuleInfo>)m_TransResult.ResultEntity;

                        if (m_lisModuleInfo.Count != 0)
                        {
                            m_TransResult.ResultEntity = m_lisModuleInfo;
                        }
                        else
                        {
                            m_TransResult.isSuccess = false;
                            m_TransResult.LogMessage = "沒有可以新增的模組,請先確認模組維護功能";
                        }

                    }
                }
            }
            catch (Exception ex)
            {
                m_TransResult.LogMessage = ex.Message;
                m_TransResult.isSuccess = false;
            }
            return m_TransResult;
        }


        public SysEntity.TransResult GetGroupInfo(SysEntity.Employee p_EmployeeEntity,string MenuID, string ActType)
        {
            SysEntity.TransResult m_TransResult = new SysEntity.TransResult();
            try
            {
                string m_Condition = " in ";
                if (ActType == "add" )
                {
                    m_Condition = " not in";
                }
                using (SqlCommand m_cmd = new SqlCommand(" SELECT [GroupDesc],[GroupID]  FROM [dbo].[View_GroupID] where [GroupID] " + m_Condition + " (SELECT  [GroupID] from [dbo].[AuthMenu] where MenuID=@MenuID) "))
                {
                    var m_Params = new
                    {
                        MenuID = MenuID
                    };

                    m_TransResult = g_DC.GetData<SysEntity.GroupInfo>(p_EmployeeEntity, m_cmd, m_Params);
                    if (m_TransResult.isSuccess)
                    {
                        List<SysEntity.GroupInfo> m_lisGroupInfo = (List<SysEntity.GroupInfo>)m_TransResult.ResultEntity;

                        if (m_lisGroupInfo.Count != 0)
                        {
                            m_TransResult.ResultEntity = m_lisGroupInfo;
                        }
                        else
                        {
                            m_TransResult.isSuccess = false;
                            m_TransResult.LogMessage = "沒有可以新增的角色,請先確認角色維護功能";
                        }

                    }
                }
            }
            catch (Exception ex)
            {
                m_TransResult.LogMessage = ex.Message;
                m_TransResult.isSuccess = false;
            }
            return m_TransResult;
        }


        public SysEntity.TransResult MaintainAuthMenu(SysEntity.Employee p_EmployeeEntity, string ActType, AuthMenu objAuthMenu)
        {
            SysEntity.TransResult m_TransResult = new SysEntity.TransResult();
            try
            {
                string m_SQL = "";
                using (SqlCommand m_cmd = new SqlCommand())
                {
                    if (ActType == "Del")
                    {
                        m_SQL = "Delete [dbo].[AuthMenu] Where [MenuID]=@MenuID and [GroupID]=@GroupID";

                        m_cmd.Parameters.Add(new SqlParameter("MenuID", objAuthMenu.MenuID));
                        m_cmd.Parameters.Add(new SqlParameter("GroupID", objAuthMenu.GroupID));
                        m_cmd.CommandText = m_SQL;
                        m_TransResult = g_DC.ExecuteNonQuery(p_EmployeeEntity, m_cmd);

                    }
                    else if (ActType == "add")
                    {
                        Dictionary<string, object> dictionary = JsonConvert.DeserializeObject<Dictionary<string, object>>(JsonConvert.SerializeObject(objAuthMenu));
                        m_TransResult = g_DC.InsertIntoByEntity<SysEntity.AuthMenu>(p_EmployeeEntity, "AuthMenu", dictionary);

                    }
                }
            }
            catch (Exception ex)
            {
                m_TransResult.LogMessage = ex.Message;
                m_TransResult.isSuccess = false;
            }
            return m_TransResult;
        }



        public SysEntity.TransResult MaintainMainMenu(SysEntity.Employee p_EmployeeEntity, string ActType, MainMenu objMainMenu)
        {
            SysEntity.TransResult m_TransResult = new SysEntity.TransResult();
            try
            {
                string m_SQL = "";
                using (SqlCommand m_cmd = new SqlCommand())
                {
                    string SqlForm = "";


                    if (ActType == "upd")
                    {
                        if (!string.IsNullOrEmpty(objMainMenu.FormID))
                        {
                            SqlForm = " ,FormID=@FormID ";
                            m_cmd.Parameters.Add(new SqlParameter("FormID", objMainMenu.FormID));
                        }
                        m_SQL = "Update [dbo].[MainMenu]  set MenuIndex =@MenuIndex " + SqlForm + " ,ModifyUser=@ModifyUser , ModifyDate=(Convert(varchar(10),Getdate(),111)), ModifyTime=(Convert(varchar(8),Getdate(),108)) ";
                        m_cmd.Parameters.Add(new SqlParameter("MenuIndex", objMainMenu.MenuIndex));
                        m_cmd.Parameters.Add(new SqlParameter("ModifyUser", p_EmployeeEntity.WorkID));
                        if (!string.IsNullOrEmpty(objMainMenu.MenuID))
                        {
                            m_SQL += " where [MenuID]=@MenuID ";
                            m_cmd.Parameters.Add(new SqlParameter("MenuID", objMainMenu.MenuID));
                            m_cmd.CommandText = m_SQL;
                            m_TransResult = g_DC.ExecuteNonQuery(p_EmployeeEntity, m_cmd);
                        }
                        else
                        {
                            m_TransResult.LogMessage = "MenuID 不可為空!!!";
                            m_TransResult.isSuccess = false;
                        }
                    }
                    else if (ActType == "Del")
                    {
                        m_SQL = "Delete [dbo].[MainMenu] Where [MenuID]=@MenuID";

                        m_cmd.Parameters.Add(new SqlParameter("MenuID", objMainMenu.MenuID));
                        m_cmd.CommandText = m_SQL;
                        m_TransResult = g_DC.ExecuteNonQuery(p_EmployeeEntity, m_cmd);

                    }
                    else if (ActType == "add")
                    {
                        Dictionary<string, object> dictionary = JsonConvert.DeserializeObject<Dictionary<string, object>>(JsonConvert.SerializeObject(objMainMenu));
                        m_TransResult = g_DC.InsertIntoByEntity<SysEntity.MainMenu>(p_EmployeeEntity, "MainMenu", dictionary);

                    }
                }
            }
            catch (Exception ex)
            {
                m_TransResult.LogMessage = ex.Message;
                m_TransResult.isSuccess = false;
            }
            return m_TransResult;
        }



        public SysEntity.TransResult GetMainModle(SysEntity.Employee p_EmployeeEntity, string p_Site)
        {
            SysEntity.TransResult m_TransResult = new SysEntity.TransResult();
            try
            {
                DataTable m_dt = new DataTable();
                string m_SQL = " SELECT [MenuID],[SiteID],[MenuType],[MenuParentID],[MenuIndex],[MenuPathParam],I.[IconCSS] ,I.ModuleName MenuDesc" +
                    " FROM [dbo].[MainMenu] M  " +
                    " left join ModuleInfo I on M.ModuleKey=I.ModuleKey  " +
                    " where  [MenuID] in ( SELECT MenuID FROM dbo.AuthMenu where GroupID in (SELECT  GroupID FROM dbo.EmployeeGroup where WorkID=@WorkID) ) " +
                    " and  MenuType in('M') and SiteID=@SiteID order by MenuIndex ";

                using (SqlCommand m_cmd = new SqlCommand(m_SQL))
                {
                    m_cmd.Parameters.Add(new SqlParameter("SiteID", p_Site));
                    m_cmd.Parameters.Add(new SqlParameter("WorkID", p_EmployeeEntity.WorkID));
                    m_TransResult = g_DC.GetDataTable(p_EmployeeEntity, m_cmd);
                    if (m_TransResult.isSuccess)
                    {
                        m_dt = (DataTable)m_TransResult.ResultEntity;
                    }
                }
            }
            catch (Exception ex)
            {
                m_TransResult.isSuccess = false;
                m_TransResult.LogMessage = ex.Message;

            }
            return m_TransResult;
        }

        public SysEntity.TransResult GetMainModleBySite(SysEntity.Employee p_EmployeeEntity, string p_Site)
        {
            SysEntity.TransResult m_TransResult = new SysEntity.TransResult();
            try
            {
                DataTable m_dt = new DataTable();
                string m_SQL = " SELECT [MenuID],M.ModuleKey,[SiteID],[MenuType],[MenuParentID],[MenuIndex],[MenuPathParam],I.[IconCSS] ,I.ModuleName MenuDesc" +
                    " FROM [dbo].[MainMenu] M  " +
                    " left join ModuleInfo I on M.ModuleKey=I.ModuleKey  " +
                    " where   MenuType in('M') and SiteID=@SiteID order by MenuIndex ";

                using (SqlCommand m_cmd = new SqlCommand(m_SQL))
                {
                    m_cmd.Parameters.Add(new SqlParameter("SiteID", p_Site));
                    m_TransResult = g_DC.GetDataTable(p_EmployeeEntity, m_cmd);
                    if (m_TransResult.isSuccess)
                    {
                        m_dt = (DataTable)m_TransResult.ResultEntity;
                    }
                }
            }
            catch (Exception ex)
            {
                m_TransResult.isSuccess = false;
                m_TransResult.LogMessage = ex.Message;

            }
            return m_TransResult;
        }

        public SysEntity.TransResult GetMenuParentID(SysEntity.Employee p_EmployeeEntity, string p_MenuID)
        {
            SysEntity.TransResult m_TransResult = new SysEntity.TransResult();
            try
            {
                DataTable m_dt = new DataTable();
                string m_SQL = " select[MenuID],[ModuleKey],[SiteID],[MenuType],[MenuParentID],[MenuIndex],[MenuDesc],[MenuPath],[MenuPathParam] from ( ";
                m_SQL += "  SELECT [MenuID],  M.ModuleKey,[SiteID],[MenuType],[MenuParentID],[MenuIndex],[Description] MenuDesc, F.[Path] MenuPath,[MenuPathParam] " +
                    "FROM [dbo].[MainMenu] M Left Join [dbo].[FormInfo] F  on M.FormID=F.FormID where  MenuParentID=@MenuParentID  ";
               
                m_SQL += " ) M  order by MenuIndex ";

                using (SqlCommand m_cmd = new SqlCommand(m_SQL))
                {
                    m_cmd.Parameters.Add(new SqlParameter("MenuParentID", p_MenuID));

                    m_TransResult = g_DC.GetDataTable(p_EmployeeEntity, m_cmd);
                    if (m_TransResult.isSuccess)
                    {
                        m_dt = (DataTable)m_TransResult.ResultEntity;
                    }
                }
            }
            catch (Exception ex)
            {
                m_TransResult.isSuccess = false;
                m_TransResult.LogMessage = ex.Message;

            }
            return m_TransResult;
        }

        public SysEntity.TransResult GetMenuByMenuID(SysEntity.Employee p_EmployeeEntity, string p_MenuID)
        {
            SysEntity.TransResult m_TransResult = new SysEntity.TransResult();
            try
            {
                DataTable m_dt = new DataTable();
                string m_SQL = " select[MenuID],[ModuleKey],[SiteID],[MenuType],[MenuParentID],[MenuIndex],[MenuDesc],[MenuPath],[MenuPathParam] from ( ";
                m_SQL += "  SELECT [MenuID], M.ModuleKey,[SiteID],[MenuType],[MenuParentID],[MenuIndex],[Description] MenuDesc, F.[Path] MenuPath,[MenuPathParam] " +
                    "FROM [dbo].[MainMenu] M Left Join [dbo].[FormInfo] F  on M.FormID=F.FormID where  MenuParentID=@MenuParentID and [AuthFlag]='N' ";
                m_SQL += " union all ";
                m_SQL += " SELECT [MenuID], M.ModuleKey,[SiteID],[MenuType],[MenuParentID],[MenuIndex],[Description] MenuDesc, F.[Path] MenuPath,[MenuPathParam]  " +
                    "FROM [dbo].[MainMenu] M Left Join [dbo].[FormInfo] F  on M.FormID=F.FormID where [AuthFlag]='Y' ";
                m_SQL += " and [MenuID] in(SELECT [MenuID] FROM [dbo].[AuthMenu] where MenuParentID=@MenuParentID and  [GroupID] in(select GroupID from [dbo].[EmployeeGroup] where WorkID=@WorkID))) M  order by MenuIndex ";

                using (SqlCommand m_cmd = new SqlCommand(m_SQL))
                {
                    m_cmd.Parameters.Add(new SqlParameter("MenuParentID", p_MenuID));
                    m_cmd.Parameters.Add(new SqlParameter("WorkID", p_EmployeeEntity.WorkID));

                    m_TransResult = g_DC.GetDataTable(p_EmployeeEntity, m_cmd);
                    if (m_TransResult.isSuccess)
                    {
                        m_dt = (DataTable)m_TransResult.ResultEntity;
                    }
                }
            }
            catch (Exception ex)
            {
                m_TransResult.isSuccess = false;
                m_TransResult.LogMessage = ex.Message;

            }
            return m_TransResult;
        }

        #endregion

    

    }
}
