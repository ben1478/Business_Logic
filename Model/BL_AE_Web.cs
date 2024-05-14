using Business_Logic.Entity;
using NPOI.OpenXmlFormats.Wordprocessing;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static Business_Logic.Entity.SysEntity;

namespace Business_Logic.Model
{

   


    public class BL_AE_Web
    {

        public class AE_Msg
        {
            public String Msg_cknum { get; set; }
            public String Msg_source { get; set; }
            public String Msg_kind { get; set; }
            public String Msg_show_date { get; set; }
            public String Msg_title { get; set; }
            public String Msg_note { get; set; }
            public String Msg_to_num { get; set; }
            public String add_num { get; set; }
            public String add_ip { get; set; }
        }

        public class FR_info
        {
            public String FR_kind { get; set; }
            public String FR_sign_type { get; set; }
            public String FR_step_now { get; set; }
            public String FR_sign_note { get; set; }
            
        }

        DataComponent g_DC = new DataComponent();

        public static FunctionHandler g_FH = new FunctionHandler();
        public SysEntity.TransResult GetRestListGrid(SysEntity.Employee p_EmployeeEntity, Dictionary<string, Object> p_dicRec, string p_PageIndex, string p_PageSize, Dictionary<string, Object> p_Param)
        {
            SysEntity.TransResult m_TransResult = new SysEntity.TransResult();
            Dictionary<string, Object> m_objBetween = new Dictionary<string, Object>();
            string[] m_LikeProp = { };
            p_dicRec.Add("Sign_num", p_EmployeeEntity.WorkID);
            ArrayList m_alGrid = new ArrayList();
            try
            {
                string m_SQL = "  SELECT R.FR_id,R.FR_cknum,M.U_name FR_U_name ,FR_kind_show,  convert(varchar(16), R.FR_date_begin, 121) FR_date_begin , " +
                               "  convert(varchar(16), R.FR_date_begin, 121) +'<br>'+convert(varchar(16), R.FR_date_end, 121) FR_date ,R.FR_total_hour,FR_note " +
                               "  ,case when FR_step_now ='1' then FR_step_01_num when FR_step_now ='2' then FR_step_02_num " +
                               "  when FR_step_now ='3' then FR_step_03_num  end Sign_num     " +
                               " ,case when FR_step_now ='1' then '代理人' when FR_step_now ='2' then '直屬主管' when FR_step_now ='3' then '單位主管'  end sign_type" +
                               " FROM Flow_rest R  Left join User_M M on R.FR_U_num=M.U_num       " +
                               "  Left join  " +
                               "  ( " +
                               "        SELECT item_D_name FR_kind_show,item_D_code FROM Item_list   " +
                               "        WHERE item_M_code = 'FR_kind' AND item_D_type='Y' AND del_tag='0'   " +
                               "  ) I on R.FR_kind=I.item_D_code  " +
                               "  WHERE R.del_tag = '0' AND [cancel_date] IS NULL  " +
                               "  and FR_step_now in ('1','2','3')  and 'FSIGN001' = case when FR_step_now ='1' then FR_step_01_sign  when FR_step_now ='2' then FR_step_02_sign  when FR_step_now ='3' then FR_step_03_sign end " +
                               "  AND FR_date_begin between  DATEADD(month, -2, GETDATE()) AND  DATEADD(month, 3, GETDATE()) ";
                              
                using (SqlCommand m_cmd = g_DC.GetQueryCommand(p_EmployeeEntity, m_SQL, p_dicRec, m_objBetween, m_LikeProp))
                {
                    m_TransResult = g_DC.GetDataTableByGrid(p_EmployeeEntity, m_cmd, p_Param, p_PageSize, p_PageIndex, "AE_DB");
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

        public SysEntity.TransResult ChkErrCount(string p_Confirm_ip, string p_GUID, string p_chk_num)
        {
            SysEntity.TransResult m_TransResult = new SysEntity.TransResult();
            try
            {
                SysEntity.Employee m_EmployeeEntity = new SysEntity.Employee();
                m_EmployeeEntity.WorkID = "ChkTokenByGUID";
                string m_SQL = " SELECT * FROM dbo.AE_WebToken where Effect_time > getdate() and Confirm_ip=@Confirm_ip and [GUID]=@GUID  ";
                using (SqlCommand m_cmd = new SqlCommand(m_SQL))
                {
                    m_cmd.Parameters.Add(new SqlParameter("Confirm_ip", p_Confirm_ip));
                    m_cmd.Parameters.Add(new SqlParameter("GUID", p_GUID));
                    m_TransResult = g_DC.GetDataTable(m_EmployeeEntity, m_cmd, "AE_DB");
                    if (m_TransResult.isSuccess)
                    {
                        string m_updSQL = "";
                        DataTable dtToken = (DataTable)m_TransResult.ResultEntity;
                        if (dtToken.Rows.Count > 0)
                        {
                            foreach (DataRow dr in dtToken.Rows)
                            {
                                string m_ErrCount = dr["ErrCount"].ToString();
                                string m_chk_num = dr["chk_num"].ToString();
                                if (p_chk_num == m_chk_num)
                                {
                                    if (Convert.ToInt16(m_ErrCount) > 2)
                                    {
                                        m_TransResult.isSuccess = false;
                                        m_TransResult.LogMessage = "帳號密碼錯誤超過3次，連結已失效!!!";
                                        return m_TransResult;
                                    }
                                    else
                                    {
                                        using (SqlCommand m_updcmd = new SqlCommand())
                                        {
                                            m_updSQL = "Update [dbo].[AE_WebToken]  set ErrCount=ErrCount+1 where [GUID]=@GUID ";
                                            m_updcmd.Parameters.Add(new SqlParameter("GUID", p_GUID));
                                            m_updcmd.CommandText = m_updSQL;
                                            m_TransResult = g_DC.ExecuteNonQuery(m_EmployeeEntity, m_updcmd, "AE_DB");
                                        }
                                        
                                    }
                                }
                                else
                                {
                                    m_TransResult.isSuccess = false;
                                    m_TransResult.LogMessage = "帳號非本次開放帳號請確認";
                                    return m_TransResult;
                                }
                                
                            }
                        }
                        else
                        {
                            m_TransResult.isSuccess = false;
                            m_TransResult.LogMessage = "連結已逾時，請重新申請連結!!!";
                            return m_TransResult;
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

        public SysEntity.TransResult ChkLogin(string IP, string WorkID, string PW, string GUID)
        {
            SysEntity.TransResult m_TransResult = new SysEntity.TransResult();
            try
            {
                SysEntity.Employee m_EmployeeEntity = new SysEntity.Employee();
                m_EmployeeEntity.WorkID = "AELoginChk";

                //登入寫入Log
                Dictionary<string, Object> m_Login_Log = new Dictionary<string, object>();
                m_Login_Log.Add("RequestIP", IP);
                m_Login_Log.Add("WorkID", WorkID);
                m_TransResult = ChkErrCount(IP, GUID, WorkID);
                if (m_TransResult.isSuccess)
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
                                using (SqlCommand m_updcmd = new SqlCommand())
                                {
                                   string m_updSQL = "Update [dbo].[AE_WebToken]  set isConfirm=@isConfirm ,Confirm_date=getdate() where [GUID]=@GUID ";
                                    m_updcmd.Parameters.Add(new SqlParameter("GUID", GUID));
                                    m_updcmd.Parameters.Add(new SqlParameter("isConfirm", "Y"));
                                    m_updcmd.CommandText = m_updSQL;
                                    m_TransResult = g_DC.ExecuteNonQuery(m_EmployeeEntity, m_updcmd, "AE_DB");
                                }

                                m_Login_Log.Add("IsPass", "Y");
                                SysEntity.Employee m_ResultEmployee = new SysEntity.Employee();
                                foreach (DataRow dr in dtEmployee.Rows)
                                {
                                    m_ResultEmployee.CurrentLanguage = "TW";
                                    m_ResultEmployee.WorkID = WorkID;
                                    m_ResultEmployee.DisplayName = dr["U_name"].ToString();
                                    m_ResultEmployee.CompanyCode = "AE";
                                    m_ResultEmployee.GroupID = "AE_DB";
                                    m_ResultEmployee.Remark = "AE";
                                    m_ResultEmployee.UserIP = IP;
                                }
                                m_TransResult.ResultEntity = (SysEntity.Employee)m_ResultEmployee;
                            }
                            else
                            {
                                m_Login_Log.Add("IsPass", "N");
                                m_TransResult.isSuccess = false;
                                m_TransResult.LogMessage = "輸入帳號密碼錯誤!!";
                            }
                            g_DC.InsertIntoByEntity<SysEntity.Login_Log>(m_EmployeeEntity, "AE_Login_Log", m_Login_Log, false);
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
        /// 檢查GUID及Confirm_ip
        /// </summary>
        /// <param name="p_GUID"></param>
        /// <param name="p_Confirm_ip"></param>
        /// <returns></returns>
        public SysEntity.TransResult UpdTokenByGUID(string p_GUID, string p_Confirm_ip, string p_chk_num)
        {
            SysEntity.TransResult m_TransResult = new SysEntity.TransResult();
            try
            {
                SysEntity.Employee m_EmployeeEntity = new SysEntity.Employee();
                m_EmployeeEntity.WorkID = "ChkTokenByGUID";
                string m_SQL = " SELECT * FROM dbo.AE_WebToken where Effect_time > getdate() and isConfirm='N' and [GUID]=@GUID  ";
                using (SqlCommand m_cmd = new SqlCommand(m_SQL))
                {
                    m_cmd.Parameters.Add(new SqlParameter("GUID", p_GUID));
                    m_TransResult = g_DC.GetDataTable(m_EmployeeEntity, m_cmd, "AE_DB");
                    if (m_TransResult.isSuccess)
                    {

                        string m_updSQL = "";
                        DataTable dtToken = (DataTable)m_TransResult.ResultEntity;
                        if (dtToken.Rows.Count > 0)
                        {
                            foreach (DataRow dr in dtToken.Rows)
                            {
                                string m_Confirm_ip = dr["Confirm_ip"].ToString();
                                string m_chk_num = dr["chk_num"].ToString();
                                if (p_chk_num != m_chk_num)
                                {
                                    m_TransResult.isSuccess = false;
                                    m_TransResult.LogMessage = "帳號非本次開放帳號請確認";
                                    return m_TransResult;
                                }
                                if (m_Confirm_ip == "")
                                {
                                    using (SqlCommand m_updcmd = new SqlCommand())
                                    {
                                        m_updSQL = "Update [dbo].[AE_WebToken]  set Confirm_ip=@Confirm_ip where [GUID]=@GUID ";
                                        m_updcmd.Parameters.Add(new SqlParameter("GUID", p_GUID));
                                        m_updcmd.Parameters.Add(new SqlParameter("Confirm_ip", p_Confirm_ip));
                                        m_updcmd.CommandText = m_updSQL;
                                        m_TransResult = g_DC.ExecuteNonQuery(m_EmployeeEntity, m_updcmd, "AE_DB");
                                    }

                                }
                                else
                                {
                                    if (m_Confirm_ip != p_Confirm_ip)
                                    {
                                        m_TransResult.isSuccess = false;
                                        m_TransResult.LogMessage = "連結已失效!!!";
                                        return m_TransResult;
                                    }
                                }
                            }
                        }
                        else
                        {
                            m_TransResult.isSuccess = false;
                            m_TransResult.LogMessage = "連結已逾時，請重新申請連結!!!";
                            return m_TransResult;
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

        public SysEntity.TransResult ChkEffectToken(string p_GUID )
        {
            SysEntity.TransResult m_TransResult = new SysEntity.TransResult();
            try
            {
                SysEntity.Employee m_EmployeeEntity = new SysEntity.Employee();
                string m_SQL = " SELECT * FROM dbo.AE_WebToken where Effect_time > getdate() and isConfirm=@isConfirm  and errCount<3 and GUID=@GUID ";
                using (SqlCommand m_cmd = new SqlCommand(m_SQL))
                {
                    m_cmd.Parameters.Add(new SqlParameter("GUID", p_GUID));
                    m_cmd.Parameters.Add(new SqlParameter("isConfirm", "N"));
                    m_TransResult = g_DC.GetDataTable(m_EmployeeEntity, m_cmd, "AE_DB");
                    if (m_TransResult.isSuccess)
                    {
                        DataTable dtEmployee = (DataTable)m_TransResult.ResultEntity;
                        if (dtEmployee.Rows.Count > 0)
                        {
                            m_TransResult.isSuccess = true;
                        }
                        else
                        {
                            m_TransResult.isSuccess = false;
                            m_TransResult.LogMessage = "連結已失效!!!";
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

        public SysEntity.TransResult AE_Approve(SysEntity.Employee p_EmployeeEntity, string[] p_arrFR_id)
        {
            SysEntity.TransResult m_TransResult = new SysEntity.TransResult();
            try
            {
                List<Business_Logic.Entity.SysEntity.CommandEntity> m_MultiCommandEntity = new List<SysEntity.CommandEntity>();
               

                foreach (string m_FR_ID in p_arrFR_id)
                {
                    SysEntity.CommandEntity m_CommandEntity = new SysEntity.CommandEntity();
                    m_TransResult = GetFlow_rest(p_EmployeeEntity, m_FR_ID, "");
                    DataTable m_FlowTable = (DataTable)m_TransResult.ResultEntity;
                    foreach (DataRow dr in m_FlowTable.Rows)
                    {
                        string m_msg_num = "";
                        string m_FR_u_num = dr["FR_u_num"].ToString();
                        string m_FR_step_now_upd = "";
                        string m_FR_sign_type = "FSIGN001";
                        string m_FR_step_now = dr["FR_step_now"].ToString();
                        string m_FR_kind = dr["FR_kind"].ToString();
                        string m_FR_step_03_type = dr["FR_step_03_type"].ToString();
                        string m_FR_step_02_num = dr["FR_step_02_num"].ToString();
                        string m_FR_step_03_num = dr["FR_step_03_num"].ToString();
                        List<System.Data.SqlClient.SqlParameter> SqlParameters = new List<SqlParameter>();
                        string m_SQL = "update  [dbo].[Flow_rest]  set FR_sign_type=@FR_sign_type,edit_num=@edit_num,edit_date=SYSDATETIME(),edit_ip=@edit_ip  ";
                        m_SQL += ",FR_step_now=@FR_step_now";
                        switch (m_FR_step_now)
                        {
                            case "1":
                                m_SQL += " , FR_step_01_type=@FR_step_type,FR_step_01_sign=@FR_step_sign, FR_step_01_date=SYSDATETIME()    ";
                                m_FR_step_now_upd = "2";
                                m_msg_num = m_FR_step_02_num;
                                break;
                            case "2":
                                m_SQL += " , FR_step_02_type=@FR_step_type,FR_step_02_sign=@FR_step_sign , FR_step_02_date=SYSDATETIME()    ";
                                if (m_FR_step_03_type == "FSTEP001")
                                {
                                    m_msg_num = m_FR_step_03_num;
                                    m_FR_step_now_upd = "3";
                                }
                                else
                                {
                                    //'公出 FRK016 忘打卡 FRK017
                                    if (m_FR_kind == "FRK016" || m_FR_kind == "FRK017")
                                    {
                                        m_FR_sign_type = "FSIGN002";
                                        m_FR_step_now_upd = "0";
                                    }
                                    else
                                    {
                                        m_FR_step_now_upd = "9";
                                    }
                                }
                                break;
                            case "3":
                                m_SQL += " , FR_step_03_type=@FR_step_type,FR_step_03_sign=@FR_step_sign , FR_step_03_date=SYSDATETIME()    ";
                                m_FR_step_now_upd = "9";
                                break;
                        }
                        
                        object m_Params = new
                        {
                            FR_sign_type = m_FR_sign_type,                         
                            edit_num = p_EmployeeEntity.WorkID,
                            edit_ip = p_EmployeeEntity.UserIP,
                            FR_step_now = m_FR_step_now_upd,
                            FR_step_sign = "FSIGN002",
                            FR_step_type= "FSTEP002",
                            FR_ID= m_FR_ID
                        };
                        m_SQL += " where FR_ID=@FR_ID  ";
                        m_CommandEntity.SQLCommand = m_SQL;
                        m_CommandEntity.Parameters = m_Params;
                        m_MultiCommandEntity.Add(m_CommandEntity);
                        Dictionary<string, object> m_disMsg = new Dictionary<string, object>();

                        m_disMsg["add_ip"] = p_EmployeeEntity.UserIP;
                        m_disMsg["Msg_cknum"] = DateTime.Now.ToString("yyyyMMddHHmmssfffffff");
                        m_disMsg["Msg_source"] = "OUTsys";
                        m_disMsg["Msg_kind"] = "MSGK0004";
                        m_disMsg["Msg_to_num"] = m_msg_num;
                        m_disMsg["Msg_title"] = "請假單簽核通知,請前往處理!!";
                        m_disMsg["Msg_show_date"] = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff");
                        m_disMsg["add_num"] = m_FR_u_num;
                        m_MultiCommandEntity.Add(g_DC.GetInsertCommand(p_EmployeeEntity, "Msg", m_disMsg));
                    }
                }
                m_TransResult = g_DC.ExecuteCommandEntitys(p_EmployeeEntity, m_MultiCommandEntity,"AE_DB");
                if (m_TransResult.isSuccess)
                {
                    m_TransResult.LogMessage = "審核成功!!";
                }
            }
            catch (Exception ex)
            {
                m_TransResult.LogMessage = ex.Message;
                m_TransResult.isSuccess = false;
            }
            return m_TransResult;
        }

        public SysEntity.TransResult AE_ApproveDtl(SysEntity.Employee p_EmployeeEntity, string p_FR_id , string p_FR_cknum, string p_FR_step_sign, string p_FR_note)
        {
            SysEntity.TransResult m_TransResult = new SysEntity.TransResult();
            try
            {
                List<Business_Logic.Entity.SysEntity.CommandEntity> m_MultiCommandEntity = new List<SysEntity.CommandEntity>();

                if (p_FR_id != "")
                {

                    /*  簽核狀態
                        'FSTEP001 待簽核
                        'FSTEP002 已簽核
                        'FSTEP003 免簽核

                        '簽核結果
                        'FSIGN001 未簽核
                        'FSIGN002 同意
                        'FSIGN003 不同意
                    */
                    AE_Msg m_AE_Msg = new AE_Msg();
                    m_TransResult = GetFlow_rest(p_EmployeeEntity, p_FR_id, p_FR_cknum);
                    if (m_TransResult.isSuccess)
                    {
                        SysEntity.CommandEntity m_CommandEntity = new SysEntity.CommandEntity();
                        string m_FR_step_02_num = "";
                        string m_FR_step_03_num = "";
                        string m_msg_num = "";
                        string m_FR_kind = "";
                        string m_FR_step_03_type = "";
                        string m_FR_step_now = "";
                        string m_FR_sign_type = "";
                        string m_FR_step_now_upd = "";
                        string m_SQL = "  ";
                        string m_SQL_upd = "  ";
                        string m_MSG = "";
                        string m_FR_u_num = "";
                        DataTable m_FlowTable = (DataTable)m_TransResult.ResultEntity;
                        foreach (DataRow dr in m_FlowTable.Rows)
                        {
                            m_FR_u_num = dr["FR_u_num"].ToString();
                            m_FR_step_now = dr["FR_step_now"].ToString();
                            m_FR_kind = dr["FR_kind"].ToString();
                            m_FR_step_03_type = dr["FR_step_03_type"].ToString();
                            m_FR_step_02_num = dr["FR_step_02_num"].ToString();
                            m_FR_step_03_num = dr["FR_step_03_num"].ToString();
                        }
                        m_SQL = "update  [dbo].[Flow_rest]  set FR_sign_type=@FR_sign_type,edit_num=@edit_num,edit_date=SYSDATETIME(),edit_ip=@edit_ip  ";
                        if (p_FR_step_sign == "FSIGN002")
                        {
                           
                            m_FR_sign_type = "FSIGN001";
                            m_SQL_upd += ",FR_step_now=@FR_step_now";
                        }
                        else
                        {
                            m_FR_sign_type = "FSIGN003";
                        }
                        switch (m_FR_step_now)
                        {
                            case "1":
                                m_SQL_upd += " , FR_step_01_type=@FR_step_type,FR_step_01_sign=@FR_step_sign,FR_step_01_note=@FR_note , FR_step_01_date=SYSDATETIME()    ";
                                if (p_FR_step_sign == "FSIGN002")
                                {
                                    m_msg_num = m_FR_step_02_num;
                                    m_FR_step_now_upd = "2";
                                }
                                break;
                            case "2":
                                m_SQL_upd += " , FR_step_02_type=@FR_step_type,FR_step_02_sign=@FR_step_sign,FR_step_02_note=@FR_note , FR_step_02_date=SYSDATETIME()    ";
                                if (m_FR_step_03_type == "FSTEP001")
                                {
                                    if (p_FR_step_sign == "FSIGN002")
                                    {
                                        m_msg_num = m_FR_step_03_num;
                                        m_FR_step_now_upd = "3";
                                    }
                                }
                                else
                                {
                                    //'公出 FRK016 忘打卡 FRK017
                                    if (m_FR_kind == "FRK016" || m_FR_kind == "FRK017")
                                    {
                                        m_FR_sign_type = "FSIGN002";
                                        m_FR_step_now_upd = "0";
                                    }
                                    else
                                    {
                                        m_FR_step_now_upd = "9";
                                    }
                                }
                                break;
                            case "3":
                                m_SQL_upd += " , FR_step_03_type=@FR_step_type,FR_step_03_sign=@FR_step_sign,FR_step_03_note=@FR_note , FR_step_03_date=SYSDATETIME()    ";
                                if (p_FR_step_sign == "FSIGN002")
                                {
                                    m_FR_step_now_upd = "9";
                                }
                                else
                                {
                                    m_FR_step_now_upd = "0";
                                    m_FR_sign_type = "FSIGN003";
                                }
                                break;
                        }
                        m_SQL_upd += " where  FR_id=@FR_id and FR_cknum=@FR_cknum ";
                        m_SQL += m_SQL_upd;
                        object m_Params = new
                        {
                            FR_sign_type = m_FR_sign_type,
                            FR_step_type = "FSTEP002",
                            edit_num = p_EmployeeEntity.WorkID,
                            edit_ip = p_EmployeeEntity.UserIP,
                            FR_step_now = m_FR_step_now_upd,
                            FR_step_sign = p_FR_step_sign,
                            FR_note= p_FR_note,
                            FR_ID = p_FR_id,
                            FR_cknum = p_FR_cknum
                        };

                        m_CommandEntity.SQLCommand = m_SQL;
                        m_CommandEntity.Parameters = m_Params;
                        m_MultiCommandEntity.Add(m_CommandEntity);
                        Dictionary<string, Object> m_disMsg = new Dictionary<string, object>();
                         
                        m_disMsg["add_ip"] = p_EmployeeEntity.UserIP;
                        m_disMsg["Msg_cknum"]= DateTime.Now.ToString("yyyyMMddHHmmssfffffff");
                        m_disMsg["Msg_source"] = "sys";
                        m_disMsg["Msg_kind"] = "MSGK0004";
                        m_disMsg["Msg_to_num"] = m_msg_num;
                        m_disMsg["Msg_title"] = "請假單簽核通知,請前往處理!!";
                        m_disMsg["Msg_show_date"] = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff");
                        m_disMsg["add_num"] = m_FR_u_num;
                        m_MultiCommandEntity.Add(g_DC.GetInsertCommand(p_EmployeeEntity, "Msg", m_disMsg));

                        m_TransResult = g_DC.ExecuteCommandEntitys(p_EmployeeEntity, m_MultiCommandEntity, "AE_DB");
                        if (m_TransResult.isSuccess)
                        {
                            m_TransResult.LogMessage = "審核成功!!";
                        }
                    }

                }
                else
                {
                    m_TransResult.isSuccess = false;
                    m_TransResult.LogMessage = "p_FR_id is null";
                }
            }
            catch (Exception ex)
            {
                m_TransResult.LogMessage = ex.Message;
                m_TransResult.isSuccess = false;
            }
            return m_TransResult;
        }

        public SysEntity.TransResult GetFlow_rest(SysEntity.Employee p_EmployeeEntity,string p_FR_id, string p_FR_cknum)
        {
            SysEntity.TransResult m_TransResult = new SysEntity.TransResult();
            try
            {
               
                string m_SQL = " SELECT * FROM dbo.Flow_rest where  FR_id=@FR_id   ";
                if (p_FR_cknum != "")
                {
                    m_SQL += "  and  FR_cknum=@FR_cknum ";
                }
                using (SqlCommand m_cmd = new SqlCommand(m_SQL))
                {
                    m_cmd.Parameters.Add(new SqlParameter("FR_id", p_FR_id));
                    if (p_FR_cknum != "")
                    {
                        m_cmd.Parameters.Add(new SqlParameter("FR_cknum", p_FR_cknum));
                    }
                    m_TransResult = g_DC.GetDataTable(p_EmployeeEntity, m_cmd, "AE_DB");
                    if (m_TransResult.isSuccess)
                    {
                        DataTable dtEmployee = (DataTable)m_TransResult.ResultEntity;
                        if (dtEmployee.Rows.Count > 0)
                        {
                            m_TransResult.isSuccess = true;
                        }
                        else
                        {
                            m_TransResult.isSuccess = false;
                            m_TransResult.LogMessage = "查無資料!!!";
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


    }
}
