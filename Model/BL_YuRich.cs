using Business_Logic.Entity;
using Newtonsoft.Json;
using NPOI.SS.Formula.Functions;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Web.Script.Serialization;
using static Business_Logic.Entity.SysEntity;

namespace Business_Logic.Model
{
    public class BL_YuRich
    {
        DataComponent g_DC = new DataComponent();

        public static FunctionHandler g_FH = new FunctionHandler();
        /*TEST*/
        public SysEntity.TransResult GetReceiveGrid(SysEntity.Employee p_EmployeeEntity, Dictionary<string, Object> p_dicRec, string p_PageIndex, string p_PageSize, Dictionary<string, Object> p_Param)
        {
            SysEntity.TransResult m_TransResult = new SysEntity.TransResult();
            Dictionary<string, Object> m_objBetween = new Dictionary<string, Object>();
            if (p_dicRec.ContainsKey("add_date_S"))
            {
                m_objBetween["add_date_S"] = p_dicRec["add_date_S"].ToString();
                m_objBetween["add_date_E"] = p_dicRec["add_date_E"].ToString();
            }
            if (p_EmployeeEntity.GroupID != "AD01" && p_EmployeeEntity.GroupID != "AS01")
            {
                if (p_EmployeeEntity.WorkID == "YL0001")
                {
                    p_dicRec.Add("Case_Company", "YL");

                }
                else {
                    p_dicRec.Add("Add_User", p_EmployeeEntity.WorkID);
                }
            }

            ArrayList m_alGrid = new ArrayList();
            string[] m_LikeProp = { "WorkID", "DisplayName" };
            try
            {
                string m_SQL = " SELECT Case_Company, YR.form_no, FORMAT( YR.add_date, 'yyyy/MM/dd')add_date,DisplayName add_name" +
                               " ,isnull(ExamineNo,'')ExamineNo,  customer_name,          " +
                               "  case when CaseStatus ='0' then '未轉裕富-'  +( dbo.GetJsonValue(ResultJSON,'msg'))  " +
                               "  when CaseStatus ='1' then '裕富收件' when CaseStatus ='2' then '核准' " +
                               "  when CaseStatus ='3' then '婉拒'     when CaseStatus ='4' then '附條件'  " +
                               "  when CaseStatus ='5' then '待補'     when CaseStatus ='6' then '補件'  " +
                               "  when CaseStatus ='7' then '申覆'     when CaseStatus ='8' then '自退'  " +
                               "  when CaseStatus ='RE' then '申覆中'  when CaseStatus ='RS' then '補件中' " +
                               "  when CaseStatus ='RP' then '請款中' when CaseStatus ='AP' then '已撥款'  " +
                               "  end CaseStatusDesc,YR.Add_User,YR.CaseStatus,  " +
                               "  case when CaseStatus in ('1','RE','RS','RP') then 'Y'else 'N' end IsAwait " +
                               "  from tbReceive YR " +
                               "  Left join [KF_DB].[dbo].[EmployeeInfo] E on  YR.Add_User=E.WorkID  " +
                               "  Left join (select Form_No, ResultJSON from [tbAPILog] where TransactionId in (  " +
                               "  select  MAX(TransactionId)TransactionId from [tbAPILog]  " +
                               "  group by Form_No)) L on YR.form_no=L.Form_No " +
                               " where  YR.Status = '1' ";
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

        public SysEntity.TransResult GetAPILogGrid(SysEntity.Employee p_EmployeeEntity, Dictionary<string, Object> p_dicRec, string p_PageIndex, string p_PageSize, Dictionary<string, Object> p_Param)
        {
            SysEntity.TransResult m_TransResult = new SysEntity.TransResult();
            Dictionary<string, Object> m_objBetween = new Dictionary<string, Object>();
            m_objBetween["add_date_S"] = p_dicRec["add_date_S"].ToString();
            m_objBetween["add_date_E"] = p_dicRec["add_date_E"].ToString();
            

            ArrayList m_alGrid = new ArrayList();


            string[] m_LikeProp = {  };

            try
            {
                string m_SQL = " SELECT form_no, FORMAT( YR.add_date, 'yyyy/MM/dd')add_date,DisplayName add_name" +
                               " ,isnull(ExamineNo,'')ExamineNo,            " +
                               "  case when CaseStatus ='0' then '未轉裕富'  " +
                               "  when CaseStatus ='1' then '裕富收件' when CaseStatus ='2' then '核准' " +
                               "  when CaseStatus ='3' then '婉拒'     when CaseStatus ='4' then '附條件'  " +
                               "  when CaseStatus ='5' then '待補'     when CaseStatus ='6' then '補件'  " +
                               "  when CaseStatus ='7' then '申覆'     when CaseStatus ='8' then '自退'  " +
                               "  when CaseStatus ='RE' then '申覆中'  when CaseStatus ='RS' then '補件中' " +
                               "  when CaseStatus ='RP' then '請款中' when CaseStatus ='AP' then '已撥款'  " +
                               "  end CaseStatusDesc,YR.Add_User,YR.CaseStatus  " +
                               "  from tbReceive YR " +
                               "  Left join [KF_DB].[dbo].[EmployeeInfo] E on  YR.Add_User=E.WorkID  " +
                               " where  Status = '1' ";
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

        public SysEntity.TransResult GetPaymentGrid(SysEntity.Employee p_EmployeeEntity, Dictionary<string, Object> p_dicRec, string p_PageIndex, string p_PageSize, Dictionary<string, Object> p_Param)
        {
            SysEntity.TransResult m_TransResult = new SysEntity.TransResult();
            Dictionary<string, Object> m_objBetween = new Dictionary<string, Object>();

            if (p_dicRec.ContainsKey("PayDate_S"))
            {
                m_objBetween["PayDate_S"] = p_dicRec["PayDate_S"].ToString();
                m_objBetween["PayDate_E"] = p_dicRec["PayDate_E"].ToString();

            }

            string[] m_LikeProp = { };

            try
            {
                string m_SQL = " select case when APP.form_no is null then 'N' else 'Y' end IsAppr, PayDate,  SUBSTRING(Q.promoName, 1, CHARINDEX('(', Q.promoName) - 1) promoName,R.ExamineNo ExamineNo_Pay, R.customer_name,R.customer_idcard_no, " +
                               " R.car_no,QP.[instNo] ,REPLACE(CONVERT(VARCHAR(12), CONVERT(MONEY, QP.[instAmt]), 1), '.00', '')instAmt ," +
                               " REPLACE(CONVERT(VARCHAR(12), CONVERT(MONEY, Q.instCap), 1), '.00', '') instCap, R.form_no," +
                               " REPLACE(CONVERT(VARCHAR(12), CONVERT(MONEY, isnull(QC.remitAmount,0)), 1), '.00', '') remitAmount            " +
                               "  from [tbRequestPayment] RP   " +
                               "  Left Join [tbQCS] Q on RP.transactionId_qcs= Q.transactionId  " +
                               "  Left Join [tbQCS_payment] QP on Q.form_no=QP.form_no and Q.qcs_idx=QP.qcs_idx  " +
                               "  Left Join (select * from tbQCS_capitalApply where [payeeTypeName]='裕富') QC on Q.form_no=QC.form_no and Q.qcs_idx=QC.qcs_idx  " +
                               "  Left Join (select form_no,qcs_idx, CONVERT(VARCHAR(10), CAST(appropriateDate AS DATE), 111) PayDate from tbQCS_capitalApply where [payeeTypeName]<>'裕富') KF on Q.form_no=KF.form_no and Q.qcs_idx=KF.qcs_idx  " +
                               "  Left Join ( select form_no,Status, car_no,ExamineNo,customer_name,customer_idcard_no,transactionId_qcs,transactionId   " +
                               "              from tbReceive where Status='1' and CaseStatus in('AP','RP')  " +
                               "             ) R on RP.form_no=R.form_no " +
                               "   Left Join tbAppropriation APP on R.form_no=APP.form_no and R.ExamineNo=APP.ExamineNo  " +
                               "  where PayDate<>''  and R.Status is not null ";
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


        public SysEntity.TransResult GetPrePaymentGrid(SysEntity.Employee p_EmployeeEntity, Dictionary<string, Object> p_dicRec, string p_PageIndex, string p_PageSize, Dictionary<string, Object> p_Param)
        {
            SysEntity.TransResult m_TransResult = new SysEntity.TransResult();
            Dictionary<string, Object> m_objBetween = new Dictionary<string, Object>();

          

            string[] m_LikeProp = { };

            try
            {
                string m_SQL = " select " +
                      " case when R.promotion_no in ('DQ01', 'DQ02', 'DQ04', 'DQ05', 'DQ06', 'DQ07') then '新_OA機車原融'" +
                               "     when R.promotion_no in ('DQ74','DQ03') then 'OA機車原融' " +
                               " end promoName,R.ExamineNo ExamineNo_Pay, R.customer_name,R.customer_idcard_no,  " +
                               " R.car_no,R.instNo, " +
                               "  REPLACE(CONVERT(VARCHAR(12), CONVERT(MONEY, R.[instAmt]), 1), '.00', '')instAmt,  " +
                               "  REPLACE(CONVERT(VARCHAR(12), CONVERT(MONEY, R.instCap), 1), '.00', '') instCap, " +
                               "  REPLACE(CONVERT(VARCHAR(12), CONVERT(MONEY, isnull(R.remitAmount,0)), 1), '.00', '') remitAmount, R.form_no     " +
                               "  from tbReceive R   " +
                               "  where  R.Status is not null and R.CaseStatus in ('AP','RP') ";
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



        public SysEntity.TransResult GetAppropriationInfos(SysEntity.Employee p_EmployeeEntity, Dictionary<string, Object> p_Params,bool isExcel=false )
        {
            SysEntity.TransResult m_TransResult = new SysEntity.TransResult();
            try
            {


                string m_SQL = " select  case when APP.form_no is null then 'ADD' else 'UPD' end APP_STATUS,  PayDate,'湧立股份有限公司'PathName,R.ExamineNo ExamineNo_Pay,customer_idcard_no,  " +
                    "  R.customer_name,SUBSTRING(Q.promoName, 1, CHARINDEX('(', Q.promoName) - 1) promoName, Q.instCap, R.form_no," +
                    " REPLACE(CONVERT(VARCHAR(12), CONVERT(MONEY, isnull(KF.KF_remitAmount,0)), 1), '.00', '') KF_remitAmount ," +
                    " REPLACE(CONVERT(VARCHAR(12), CONVERT(MONEY, isnull(QC.YR_remitAmount,0)), 1), '.00', '') YR_remitAmount, " +
                    "  Appr_Type,HandlingFee,PathFee,CustTotAmt,RemitFee,ActualAmt,PathAmt,R.BankCode,R.BankName ,R.BankID,R.AccountID,R.AccountName ,car_no , isnull( CONVERT(VARCHAR(10), CAST(Transfer_date AS DATE), 111),'') Transfer_date          " +
                    "  from tbRequestPayment RP  " +
                    "  Left Join tbQCS Q on RP.transactionId_qcs= Q.transactionId   " +
                    "  Left Join (select remitAmount YR_remitAmount, * from tbQCS_capitalApply where payeeTypeName='裕富') QC on Q.form_no=QC.form_no and Q.qcs_idx=QC.qcs_idx   " +
                   "   Left Join (select remitAmount KF_remitAmount,form_no,qcs_idx, CONVERT(VARCHAR(10), CAST(appropriateDate AS DATE), 111) PayDate from tbQCS_capitalApply where payeeTypeName<>'裕富') KF on Q.form_no=KF.form_no and Q.qcs_idx=KF.qcs_idx   " +
                   "   Left Join ( select form_no,Status, car_no,ExamineNo,customer_name,customer_idcard_no,transactionId_qcs,transactionId,BankCode,BankName ,BankID,AccountID,AccountName   " +
                   "               from tbReceive where Status='1' and CaseStatus='AP'   " +
                   "              ) R on RP.form_no=R.form_no   " +
                   "    Left Join tbAppropriation APP on R.form_no=APP.form_no and R.ExamineNo=APP.ExamineNo   where PayDate<>'' and R.Status is not null ";
                if (isExcel)
                {
                    m_SQL += " and APP.form_no is not null  and PayDate between @PayDateS and @PayDateE";
                }
                else
                {
                    m_SQL += " and R.form_no=@form_no and R.ExamineNo=@ExamineNo_Pay  ";
                }

                using (SqlCommand m_cmd = new SqlCommand(m_SQL))
                {

                    foreach (KeyValuePair<string, Object> item in p_Params)
                    {
                        m_cmd.Parameters.Add(new SqlParameter(item.Key, item.Value));
                    }


                    DataTable m_DataTable = g_DC.GetDataTable(m_cmd, "AE_DB");
                    m_TransResult.isSuccess = true;
                    m_TransResult.ResultEntity = (m_DataTable);
                }
            }
            catch (Exception ex)
            {
                m_TransResult.LogMessage = ex.Message;
                m_TransResult.isSuccess = false;
            }
            return m_TransResult;
        }

        public SysEntity.TransResult GetApprExcelInfo(SysEntity.Employee p_EmployeeEntity, Dictionary<string, Object> p_Params)
        {
            SysEntity.TransResult m_TransResult = new SysEntity.TransResult();
            try
            {
                string m_SQL = " select " +
                    " case when R.promotion_no in ('DQ01', 'DQ02', 'DQ04', 'DQ05', 'DQ06', 'DQ07') then '新_OA機車原融'" +
                             "     when R.promotion_no in ('DQ74','DQ03') then 'OA機車原融' " +
                             " end promoName,R.ExamineNo ExamineNo_Pay, R.customer_name,R.customer_idcard_no,  " +
                             " R.car_no,R.instNo, " +
                             "  REPLACE(CONVERT(VARCHAR(12), CONVERT(MONEY, R.[instAmt]), 1), '.00', '')instAmt,  " +
                             "  REPLACE(CONVERT(VARCHAR(12), CONVERT(MONEY, R.instCap), 1), '.00', '') instCap, " +
                             "  REPLACE(CONVERT(VARCHAR(12), CONVERT(MONEY, isnull(R.remitAmount,0)), 1), '.00', '') remitAmount, R.form_no     " +
                             "  from tbReceive R   " +
                             "  where  R.Status is not null and R.CaseStatus in ('AP','RP') and R.form_no in (  ";

                Int32 ParamIdx = 0;
                foreach (KeyValuePair<string, Object> item in p_Params)
                {
                    if (ParamIdx == 0)
                    {
                        m_SQL += " @" + item.Key;
                    }
                    else
                    {
                        m_SQL += " ,@" + item.Key;
                    }
                   
                    ParamIdx++;
                }
                m_SQL += " )" ;

                using (SqlCommand m_cmd = new SqlCommand(m_SQL))
                {

                    foreach (KeyValuePair<string, Object> item in p_Params)
                    {
                        m_cmd.Parameters.Add(new SqlParameter(item.Key, item.Value));
                    }


                    DataTable m_DataTable = g_DC.GetDataTable(m_cmd, "AE_DB");
                    m_TransResult.isSuccess = true;
                    m_TransResult.ResultEntity = (m_DataTable);
                }
            }
            catch (Exception ex)
            {
                m_TransResult.LogMessage = ex.Message;
                m_TransResult.isSuccess = false;
            }
            return m_TransResult;
        }

        public SysEntity.TransResult SaveAppropriation(SysEntity.Employee p_EmployeeEntity, Dictionary<string, Object> p_Entity)
        {
            SysEntity.TransResult m_TransResult = new SysEntity.TransResult();
            try
            {
                if (p_Entity["APP_STATUS"].ToString() == "ADD")
                {
                    m_TransResult = g_DC.InsertIntoByEntity<SysEntity.tbAppropriation>(p_EmployeeEntity, "tbAppropriation", p_Entity, true, "AE_DB");
                    if (m_TransResult.isSuccess)
                    {
                        m_TransResult.LogMessage = "存檔成功";
                    }
                }
                else
                {
                    string m_SQL = "";
                    using (SqlCommand m_cmd = new SqlCommand())
                    {

                        m_SQL = "Update tbAppropriation  set Appr_Type =@Appr_Type, HandlingFee =@HandlingFee, PathFee =@PathFee,CustTotAmt=@CustTotAmt , RemitFee =@RemitFee" +
                          " , ActualAmt =@ActualAmt  , PathAmt = @PathAmt , Transfer_date =@Transfer_date, Upd_User=@Upd_User, Upd_date=Getdate() WHERE Form_no=@Form_no and ExamineNo=@ExamineNo " +
                          "  ";
                        m_cmd.Parameters.Add(new SqlParameter("Appr_Type", p_Entity["Appr_Type"].ToString()));
                        m_cmd.Parameters.Add(new SqlParameter("HandlingFee", p_Entity["HandlingFee"].ToString()));
                        m_cmd.Parameters.Add(new SqlParameter("PathFee", p_Entity["PathFee"].ToString()));
                        m_cmd.Parameters.Add(new SqlParameter("CustTotAmt", p_Entity["CustTotAmt"].ToString()));
                        m_cmd.Parameters.Add(new SqlParameter("RemitFee", p_Entity["RemitFee"].ToString()));
                        m_cmd.Parameters.Add(new SqlParameter("ActualAmt", p_Entity["ActualAmt"].ToString()));
                        m_cmd.Parameters.Add(new SqlParameter("PathAmt", p_Entity["PathAmt"].ToString()));
                        m_cmd.Parameters.Add(new SqlParameter("Upd_User", p_EmployeeEntity.WorkID));
                        m_cmd.Parameters.Add(new SqlParameter("Transfer_date", p_Entity["Transfer_date"].ToString()));
                        m_cmd.Parameters.Add(new SqlParameter("Form_no", p_Entity["Form_no"].ToString()));
                        m_cmd.Parameters.Add(new SqlParameter("ExamineNo", p_Entity["ExamineNo"].ToString()));


                        m_cmd.CommandText = m_SQL;
                        m_TransResult = g_DC.ExecuteNonQuery(p_EmployeeEntity, m_cmd, "AE_DB");
                        if (m_TransResult.isSuccess)
                        {
                            m_TransResult.LogMessage = "存檔成功";
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

        public DataTable GetAPILogInfo(string Form_No, string API_Name)
        {
            DataTable m_RtnDT = new DataTable();    
            try
            {
                string m_SQL = "SELECT TOP 1 ResultJSON  FROM tbAPILog where Form_No =@Form_No and API_Name=@API_Name   order by CallTime desc  ";
                using (SqlCommand m_cmd = new SqlCommand(m_SQL))
                {
                    m_cmd.Parameters.Add(new SqlParameter("Form_No", Form_No));
                    m_cmd.Parameters.Add(new SqlParameter("API_Name", API_Name));
                    m_RtnDT = g_DC.GetDataTable(m_cmd, "AE_DB");
                }
            }
            catch
            {
                throw;
            }
            return m_RtnDT;
        }

        public SysEntity.TransResult ChkReceiveStatus(SysEntity.Employee p_EmployeeEntity)
        {
            SysEntity.TransferJobLog m_TransferJobLog = new SysEntity.TransferJobLog();
            BL_System m_BL_System = new BL_System(); 
            m_TransferJobLog.TransferID = "ChkReceiveStatus";
            m_TransferJobLog.TransferUser = p_EmployeeEntity.WorkID;
            m_TransferJobLog.TransferStatus = "Success";
            m_TransferJobLog.TransferLog = "Success";
            m_BL_System.AddTransferJobLog(p_EmployeeEntity, m_TransferJobLog);
            SysEntity.TransResult m_TransResult = new SysEntity.TransResult();
            string m_LINE_MSG = "";
            try
            {
                m_TransResult.isSuccess = true;
                string m_SQL = " SELECT form_no,isnull(ExamineNo,'')ExamineNo,YR.CaseStatus  " +
                              "  from tbReceive YR  where  YR.Status = '1' and CaseStatus in ('1','RE','RS','RP') ";
                DataTable m_dtConcurrent = g_DC.GetDataTable(new SqlCommand(m_SQL), "AE_DB");
                foreach (DataRow dr in m_dtConcurrent.Rows)
                {
                    string CaseStatus = dr["CaseStatus"].ToString();
                    string ExamineNo = dr["ExamineNo"].ToString();
                    string form_no = dr["form_no"].ToString();
                    bool isSentLine = false;
                    /*判斷是否為請款*/
                    if (CaseStatus == "RP")
                    {
                        /*請款要先檢查狀態*/
                        m_TransResult = CallAPI("QueryAppropriation", form_no, ExamineNo, CaseStatus, "N");
                        if (m_TransResult.isSuccess && m_TransResult.LogMessage != "" && m_TransResult.LogMessage != null)
                        {
                            m_TransResult = CallAPI("QueryCaseStatus", form_no, ExamineNo, CaseStatus, "N");
                            if (m_TransResult.isSuccess && m_TransResult.LogMessage != "" && m_TransResult.LogMessage != null)
                            {
                                JavaScriptSerializer serializer = new JavaScriptSerializer();
                                dynamic m_JSONData = serializer.Deserialize<dynamic>(m_TransResult.ResultEntity.ToString());

                                DataTable m_tbOldLog = GetAPILogInfo(form_no, "QueryCaseStatus");
                                dynamic m_OldData = null;
                                foreach (DataRow row in m_tbOldLog.Rows)
                                {
                                    m_OldData = serializer.Deserialize<dynamic>(row["ResultJSON"].ToString());
                                } 
                                //比對跟上次結果是否一樣,不一樣才更新
                                bool isSame = (JsonConvert.SerializeObject(m_JSONData["objResult"]["capitalApply"]) == JsonConvert.SerializeObject(m_OldData["capitalApply"]));

                                if (!isSame)
                                {
                                    /*更新資料*/
                                    m_TransResult = CallAPI("QueryAppropriation", form_no, ExamineNo, CaseStatus, "Y");
                                    if (m_TransResult.isSuccess)
                                    {
                                        m_LINE_MSG = m_TransResult.LogMessage;
                                        m_TransResult = CallAPI("QueryCaseStatus", form_no, ExamineNo, CaseStatus, "Y");
                                        if (m_TransResult.isSuccess)
                                        {
                                            m_LINE_MSG += m_TransResult.LogMessage;
                                            isSentLine = true;
                                        }
                                        else
                                        {
                                            m_LINE_MSG = ExamineNo + ":QueryCaseStatus更新失敗!!" + m_TransResult.LogMessage;
                                        }
                                    }
                                    else
                                    {
                                        m_LINE_MSG = ExamineNo + ":QueryAppropriation更新失敗!!" + m_TransResult.LogMessage;
                                    }
                                }
                            }
                        }
                    }
                    else
                    {
                        m_TransResult = CallAPI("QueryCaseStatus", form_no, ExamineNo, CaseStatus, "N");
                        if (m_TransResult.isSuccess && m_TransResult.LogMessage != "" && m_TransResult.LogMessage != null)
                        {
                            JavaScriptSerializer serializer = new JavaScriptSerializer();
                            dynamic m_JSONData = serializer.Deserialize<dynamic>(m_TransResult.ResultEntity.ToString());

                            DataTable m_tbOldLog = GetAPILogInfo(form_no, "QueryCaseStatus");
                            dynamic m_OldData = null;
                            foreach (DataRow row in m_tbOldLog.Rows)
                            {
                                m_OldData = serializer.Deserialize<dynamic>(row["ResultJSON"].ToString());
                            }
                            bool isSame = false;
                            if (m_OldData != null)
                            {
                                isSame = (JsonConvert.SerializeObject(m_JSONData["objResult"]["reasonSuggestionDetail"]) == JsonConvert.SerializeObject(m_OldData["reasonSuggestionDetail"]));
                            }
                            //比對跟上次結果是否一樣,不一樣才更新

                            if (!isSame)
                            {
                                m_TransResult = CallAPI("QueryCaseStatus", form_no, ExamineNo, CaseStatus, "Y");
                                if (m_TransResult.isSuccess && m_TransResult.LogMessage != "" && m_TransResult.LogMessage != null)
                                {
                                    isSentLine = true;
                                    m_LINE_MSG = m_TransResult.LogMessage;
                                }
                            }
                        }
                    }
                    if (isSentLine)
                    {
                        g_FH.Send_LINE(m_LINE_MSG);
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



        public SysEntity.TransResult CallAPI(string APIName, string formNo, string ExamineNo, string CaseStatus, string isUpdDB)
        {
            SysEntity.TransResult m_TransResult = new SysEntity.TransResult();
            try
            {
                string m_MSG = "";
                m_TransResult = CallApiAndGetResponse(APIName, formNo, ExamineNo, isUpdDB).Result;
                if (m_TransResult.isSuccess)
                {
                    JavaScriptSerializer serializer = new JavaScriptSerializer();
                    dynamic data = serializer.Deserialize<dynamic>(m_TransResult.ResultEntity.ToString());
                    if (APIName == "QueryCaseStatus")
                    {
                        string result = data["resultCode"];
                        string examStatus = "";
                        if (result == "000")
                        {
                            foreach (KeyValuePair<string, Object> item in data["objResult"])
                            {
                                if (item.Key == "examStatusExplain")
                                {
                                    examStatus = item.Value.ToString();
                                }
                            }
                        }

                        switch (examStatus)
                        {
                            case "核准":
                            case "婉拒":
                            case "附條件":
                            case "待補":
                            case "自退":
                                m_MSG = "案件狀態:"+ examStatus;
                                break;
                        }
                    }
                    else
                    {
                        string result = data["resultCode"];
                        string status = "";
                        if (result == "000")
                        {
                            string code = data["objResult"]["code"].ToString();
                            if (code == "S001")
                            {
                                foreach (Dictionary<string, Object> item in data["objResult"]["appropriations"])
                                {
                                    if (item.ContainsKey("status"))
                                    {
                                        status = item["status"].ToString();
                                    }
                                }
                                if (status == "A004")
                                {
                                    m_MSG = "撥款狀態:已撥款;";
                                }
                            }
                        }
                    }
                    //不等於這4個狀態不需要更新
                    if (!(CaseStatus == "1" || CaseStatus == "RE" || CaseStatus == "RS" || CaseStatus == "RP"))
                    {
                        m_MSG = "";
                    }

                    if (m_MSG != "")
                    {
                        m_MSG = "ExamineNo:"+ ExamineNo+"-" + m_MSG;
                        m_TransResult.LogMessage = m_MSG;
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



        public async Task<SysEntity.TransResult> CallApiAndGetResponse(string APIName, string formNo, string ExamineNo , string isUpdDB)
        {
            SysEntity.TransResult m_TransResult = new SysEntity.TransResult();
            try
            {
                APIClass m_APIClass = new APIClass();
                m_APIClass.branchNo = "0001";
                m_APIClass.dealerNo = "MM09";
                m_APIClass.salesNo = "admin";
                m_APIClass.source = "22";
                m_APIClass.examineNo = ExamineNo;


                string apiUrl = "http://192.168.1.243/KF_WebAPI/" + APIName + "?Form_No=" + formNo + "&isUpdDB=" + (isUpdDB=="Y").ToString();
                using (HttpClient httpClient = new HttpClient())
                {
                    var request = new HttpRequestMessage(HttpMethod.Post, apiUrl);
                    var jsonContent = new StringContent(JsonConvert.SerializeObject(m_APIClass), Encoding.UTF8, "application/json");
                    request.Content = jsonContent;
                    var response = await httpClient.SendAsync(request);

                    if (response.IsSuccessStatusCode)
                    {
                        m_TransResult.isSuccess = true;
                        m_TransResult.ResultEntity = await response.Content.ReadAsStringAsync();
                    }
                    else
                    {
                        m_TransResult.isSuccess = false;
                        m_TransResult.LogMessage = "Failed to call WebAPI. Status code: " + response.StatusCode;
                    }
                }
            }
            catch (Exception ex)
            {
                m_TransResult.isSuccess = false;
                m_TransResult.LogMessage = "Failed to call WebAPI. Status code: " + ex.Message;
            }
            return m_TransResult;
        }

    }
}
