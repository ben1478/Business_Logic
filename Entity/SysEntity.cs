using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Collections;

namespace Business_Logic.Entity
{
    public class SysEntity
    {

        public class Login_Log
        {
            public String RequestIP { get; set; }
            public String WorkID { get; set; }
            public String IsPass { get; set; }
        }

        public class MainMenu
        {
            public String MenuID { get; set; }
            public String ModuleKey { get; set; }
            public String SiteID { get; set; }
            public String MenuType { get; set; }
            public String MenuParentID { get; set; }
            public String MenuIndex { get; set; }
            public String MenuPathParam { get; set; }
            public String AuthFlag { get; set; }
            public String MenuDesc { get; set; }
            public String FormDesc { get; set; }
            public String FormID { get; set; }


        }



        public class OwnerMaintain
        {

           
            public String Site { get; set; }//區域
            public String System1 { get; set; }//系統
            public String System3 { get; set; }//功能
            public String CaseOwner { get; set; }//CaseOwner
            public String Deputy1 { get; set; }//CaseOwner
            public String Deputy2 { get; set; }//CaseOwner
            public String AgentDateS { get; set; }//AgentDate
            public String AgentTimeS { get; set; }//AgentTime
            public String AgentDateE { get; set; }//AgentDate
            public String AgentTimeE { get; set; }//AgentTime
        }
        /// <summary>
        /// OLD EIP 待簽文件
        /// </summary>
        public class EIP_SignBox
        {

            public String WorkID { get; set; }//
            public String Desc { get; set; }//
            public String Linkdata { get; set; }//
            public String CreateUser { get; set; }//
        
        }
        public class CommandEntity
        {
            public String SQLCommand { get; set; }//
            public List<System.Data.SqlClient.SqlParameter> SqlParameters { get; set; }
            public object Parameters { get; set; }

        }


        public class BaseEntity
        {
            public String CreateUser { get; set; }//
            public String ModifyUser { get; set; }
            public String CreateTime { get; set; }
            public String ModifyTime { get; set; }
        }



        public class FormTransfer
        {
            public String TransferID { get; set; }//
            public String TransferUser { get; set; }//轉送人
            public String TransferRemark { get; set; }//轉送備註
            public Dictionary<string,string>  TransKeyValue { get; set; }//轉送Key;支援多筆
            public String CreateUser { get; set; }
            public String CreateTime { get; set; }

        }

      

        public class ViewDetail
        {
            public String SiteFormID { get; set; }
            public String ViewDetailID { get; set; }
            public String SectionID { get; set; }
            public String SectionType { get; set; }
            public String ControlType { get; set; }
            public String FieldID { get; set; }
            public String QueryID { get; set; }
            public String ControlRef { get; set; }
            public String ValidID { get; set; }
            public String Height { get; set; }
            public String Width { get; set; }
            public String EventID { get; set; }
            public String MaxLength { get; set; }
            public String ModifyKey { get; set; }
            public String DefaultValue { get; set; }
            public String QueryMethod { get; set; }
            public String CreateUser { get; set; }
            public String ModifyUser { get; set; }
            public String CreateTime { get; set; }
            public String ModifyTime { get; set; }

            
        }

        public class HyperlinkEntity
        {
            public String SiteFormID { get; set; }
            public Dictionary<string, string> KeyValues { get; set; }
        }



        public class TransferJobLog
        {
            public String TransferID { get; set; }
            public String TransferLog { get; set; }
            public String TransferStatus { get; set; }
            public String TransferUser { get; set; }
          
        }

        public class FileFolder
        {
            public String FileKey { get; set; }
            public String DocID { get; set; }
            public String SiteFormID { get; set; }
            public String FileName { get; set; }
            public String FileType { get; set; }
            public String FileIndex { get; set; }
            public String FieldID { get; set; }
            public Byte[] FileEntity { get; set; }
            public String CreateUser { get; set; }
            public String ModifyUser { get; set; }
            public String CreateDate { get; set; }
            public String ModifyDate { get; set; }
        }

        public class TransferJobSetting
        {
            /// <summary>
            /// 轉檔編號Key
            /// </summary>
            public String TransferID { get; set; }
            /// <summary>
            /// 轉檔說明
            /// </summary>
            public String TransferDesc { get; set; }
            /// <summary>
            /// 對應的EventID (Table EventSetting)
            /// </summary>
            public String TransferEventID { get; set; }
            /// <summary>
            /// 轉檔周期
            /// </summary>
            public String TransferCycle { get; set; }
            /// <summary>
            /// 轉檔單位
            /// </summary>
            public String TransferUnit { get; set; }
            /// <summary>
            /// 轉檔參數
            /// </summary>
            public String TransferParam { get; set; }
            /// <summary>
            /// 負責人(轉檔結果會發送Mail,填入工號即可)
            /// </summary>
            public String TransferOwner { get; set; }
            /// <summary>
            /// 轉檔狀態(Y:啟用;N:停用)
            /// </summary>
            public String TransferStatus { get; set; }
            public String CreateDate { get; set; }
            public String CreateUser { get; set; }
            public String ModifyUser { get; set; }
            public String ModifyDate { get; set; }
        }

        public class MultiLanguage
        {
            public String LanguageKey { get; set; }
            public String ValueTW { get; set; }
            public String ValueCN { get; set; }
            public String ValueUS { get; set; }
            public String ValueJP { get; set; }
            public String ValueBR { get; set; }
            public String CreateUser { get; set; }
            public String ModifyUser { get; set; }
            public String CreateDate { get; set; }
            public String ModifyDate { get; set; }

        }

        public class ValidatedInfo
        {
            public String ValidID { get; set; }
            public String Type { get; set; }
            public String Source { get; set; }
            public String Parameter { get; set; }
            public String Description { get; set; }
        }


        public class DocSigning
        {
            public String DocID { get; set; }
            public String SignType { get; set; }
            public String Signer { get; set; }
            public String Signed { get; set; }
            public String CounterSign { get; set; }
            public String CounterSignLevel { get; set; }
        }

        public enum DocSignType
        {
            Assign = 1,//指定簽核
            Or = 2,//多筆簽核過一及過下一關
            And = 3,//多筆簽核全過才過下一關
            Countersign = 4,//會簽
        }

        public enum DocSignAction
        {
            Appove = 1,//送審
            Transfer = 2,//轉送
            Reject = 3,//駁回
            Cancel = 4,//註銷
            CounterSign = 5,//會簽
        }

      
        /// <summary>
        /// 共用簽核類別
        /// </summary>
        public class AssignDocSigning
        {
            /// <summary>
            /// 表單號碼
            /// </summary>
            public String DocID { get; set; }
            /// <summary>
            /// 簽核類別
            /// </summary>
            public DocSignType SignType { get; set; }
            /// <summary>
            /// Next簽核人
            /// </summary>
            public String NextSigner { get; set; }
            /// <summary>
            /// 是否結案
            /// </summary>
            public Boolean isComplete { get; set; }
            /// <summary>
            /// 簽核動作
            /// </summary>
            public DocSignAction SignAction { get; set; }
            /// <summary>
            /// 轉送人
            /// </summary>
            public String Transfer { get; set; }
            /// <summary>
            /// 會簽人
            /// </summary>
            public String CounterSigner { get; set; }
            /// <summary>
            /// 會簽類別
            /// </summary>
            public DocSignType CounterSignType { get; set; }
            /// <summary>
            /// 會簽級別(會簽人要加在第幾關之後)
            /// </summary>
            public Int32 CounterSignLevel { get; set; }
            
        }
       

        public class SiteForm
        {
            public String SiteFormID { get; set; }
            public String FormID { get; set; }
            public String SiteID { get; set; }
            public String Admin { get; set; }
            public String NotesPath { get; set; }
            public String AgentName { get; set; }
            public String DocLockTime { get; set; }
            public String ValidID { get; set; }
            public String Levels { get; set; }
            public String CreateUser { get; set; }
            public String ModifyUser { get; set; }
            public String CreateDate { get; set; }
            public String ModifyDate { get; set; }
        }

        public class Form
        {
            public String FormID { get; set; }
            public String Module { get; set; }
            public String Description { get; set; }
            public String Path { get; set; }
            
            public String CreateUser { get; set; }
            public String ModifyUser { get; set; }
            public String CreateDate { get; set; }
            public String ModifyDate { get; set; }
        }

        public class ModuleInfo
        {
            public String ModuleKey { get; set; }
            public String ModuleName { get; set; }
            public String IconCSS { get; set; }
            public String ModulePagedep { get; set; }
            public String ModuleContent { get; set; }

            public String ModuleImg { get; set; }
            public String CreateUser { get; set; }
            public String ModifyUser { get; set; }
            public String CreateDate { get; set; }
            public String ModifyDate { get; set; }
        }



        public class FormInfo
        {
            public String FormID { get; set; }
            public String Path { get; set; }
            public String Description { get; set; }
          
            public String CreateUser { get; set; }
            public String ModifyUser { get; set; }
            public String CreateDate { get; set; }
            public String ModifyDate { get; set; }
        }



        public class Site
        {
            public String SiteID { get; set; }
            public String Language { get; set; }
            public String Currency { get; set; }
            public String Description { get; set; }
            public String CreateUser { get; set; }
            public String ModifyUser { get; set; }
            public String CreateDate { get; set; }
            public String ModifyDate { get; set; }
        }

        
        public class EmployeeInfo
        {

            public String WorkID { get; set; }
            public String DisplayName { get; set; }
            public String CompanyCode { get; set; }
            public String AreaName { get; set; }

            
            public String Password { get; set; }
            public String Remark { get; set; }
            public String Owenr { get; set; }
           /// <summary>
           /// 子帳號GroupID
           /// </summary>
            public String GroupID { get; set; }
            
        }


        public class AuthMenu
        {
            public String MenuID { get; set; }
            public String GroupID { get; set; }
            public String GroupDesc { get; set; }

        }

        public class GroupInfo
        {
            public String GroupID { get; set; }
            public String GroupDesc { get; set; }
        }

         
        public class Employee
        {
          
            public String CompanyCode { get; set; }
            public String WorkID { get; set; }
            public String CurrentSiteForm { get; set; }
            public String DisplayName { get; set; }
            public String CurrentLanguage { get; set; }
            public String Remark { get; set; }
            public String GroupID { get; set; }

        }



        public class PageCarousel
        {
            public String SiteID { get; set; }
            public String DocID { get; set; }
            public String CarouselType { get; set; }
            public String CarouselCaption { get; set; }
            public String CarouselDepiction { get; set; }
            public String CarouselLink { get; set; }
            public String Start_Period { get; set; }
            public String End_Period { get; set; }
            public String Status { get; set; }
            public String CreateUser { get; set; }
            public String ModifyUser { get; set; }
            public String CreateDate { get; set; }
            public String ModifyDate { get; set; }
        }
        public class EmployeeGroup
        {
            public String GroupID { get; set; }
            public String WorkID { get; set; }
            public String Group_Desc { get; set; }

        }

        public class AuthAssign
        {
            public String AssignType { get; set; }
            public String AuthID { get; set; }
            public String referID { get; set; }
        }

        public class Parameter
        {
            public String ParamID { get; set; }
            public String ParentParamID { get; set; }
            public String Text { get; set; }
            public String Value { get; set; }
            public String Description { get; set; }
            public String CreateUser { get; set; }
            public String ModifyUser { get; set; }
            public String CreateDate { get; set; }
            public String ModifyDate { get; set; }
        }


        public class QueryForm
        {
            public String QueryID { get; set; }
            public String Description { get; set; }
            public String SectionType { get; set; }
            public String ControlType { get; set; }
            public String FieldID { get; set; }
            public String ControlRef { get; set; }
            public String Height { get; set; }
            public String Width { get; set; }
            public String EventID { get; set; }
            public String MaxLength { get; set; }
            public String CreateUser { get; set; }
            public String ModifyUser { get; set; }
            public String CreateDate { get; set; }
            public String ModifyDate { get; set; }
        }

        public class EventLog
        {
            public String LogID { get; set; }
            public String LogResult { get; set; }
            public String LogMessage { get; set; }
            public String Sql { get; set; }
            public Employee Employee { get; set; }

        }

        public class FormulaEntity
        {
            public String ColumnName { get; set; }
            public String Formula { get; set; }
            public String[] Values { get; set; }

        }

        

        /// <summary>
        /// 簽核結果
        /// </summary>
        public class DocSignResult
        {   /// <summary>
            /// 是否成功
            /// </summary>
            public bool isSuccess { get; set; }
            /// <summary>
            /// 是否進入下一關
            /// </summary>
            public bool isNext { get; set; }
            /// <summary>
            /// 錯誤訊息
            /// </summary>
            public String LogMessage { get; set; }
            /// <summary>
            /// 回傳物件
            /// </summary>
            public List<SysEntity.CommandEntity> listCommandEntity { get; set; }
        }

        public class TransResult
        {
            public bool isSuccess { get; set; }
            public String LogMessage { get; set; }
            public Object ResultEntity { get; set; }
            public String GridKey { get; set; }
         }

        public class EventSetting
        {
            public String EventID { get; set; }
            public String EventModel { get; set; }
            public String EventAction { get; set; }
            public String EventDescription { get; set; }
            public String EventActionType { get; set; }
            public String EventFollow { get; set; }
            public String EventRef { get; set; }
            public String CreateUser { get; set; }
            public String ModifyUser { get; set; }
            public String CreateDate { get; set; }
            public String ModifyDate { get; set; }
            
        }


     
       

        /// <summary>
        /// 發送Email物件
        /// </summary>
        public class SmtpMail
        {   
            /// <summary>
            /// 主旨
            /// </summary>
            public String Subject { get; set; }
            public String MailTo { get; set; }
            public String MailCC { get; set; }
            public String Priority { get; set; }
            public String Body { get; set; }
            public List<Dictionary<string, Object>> Attachment { get; set; }

            
        }

        


        public class Action
        {
            public String ActionType { get; set; }
        }


        //From Danny_Chu----S
        public class QAMaster
        {
            public String QASystemSN { get; set; }
            public String sFatherTW { get; set; }
            public String sTextTW { get; set; }
            public String sValueTW { get; set; }
            public String OrderIndex { get; set; }
            public String Owner1 { get; set; }
            public String System1 { get; set; }
            public String System2 { get; set; }
            public String System3 { get; set; }
            public String System4 { get; set; }
            public String System5 { get; set; }

            public String NewQAID { get; set; }
            public String QAID { get; set; }
            public String DocID { get; set; }//for FileUpload
            public String PERNR { get; set; }
            public String User_Contact { get; set; }
            public String User_ContactALL { get; set; }
            public String User_Department { get; set; }
            public String Company { get; set; }
            public String Site { get; set; }
            public String Status { get; set; }
            public String Priority { get; set; }
            public String System { get; set; }
            public String KeyWords { get; set; }
            public String QASubject { get; set; }
            public String QADescription { get; set; }
            public String QAReply { get; set; }
            public String Case_Owner { get; set; }
            public String Case_OwnerALL { get; set; }
            public String Issue_Type { get; set; }
            public String Finish_Date { get; set; }

            public String CreateUser { get; set; }
            public String CreateDate { get; set; }
            public String CreateTime { get; set; }
            public String ModifyUser { get; set; }
            public String ModifyDate { get; set; }
            public String ModifyTime { get; set; }
        }
        //From Danny_Chu----E

        public class APIClass
        {
            /// <summary>
            /// 1.通路商編號
            /// </summary>
            public String dealerNo { get; set; }

            /// <summary>
            /// 2.據點編號
            /// </summary>
            public String branchNo { get; set; }

            /// <summary>
            /// 3.業務人員 ID
            /// </summary>
            public String salesNo { get; set; }

            /// <summary>
            /// 4.審件編號 
            /// </summary>
            public String examineNo { get; set; }

            /// <summary>
            /// 5.來源 
            /// </summary>
            public String source { get; set; }

        }

        public class reasonSuggestionDetails
        {
            public reasonSuggestionDetail[] reasonSuggestionDetail { get; set; } = { };
        }
        public class reasonSuggestionDetail
        {
            /// <summary>
            /// 原因/建議  
            /// </summary>
            public string kind { get; set; } = "";
            /// <summary>
            /// 審件狀態  
            /// </summary>
            public string explain { get; set; } = "";
            /// <summary>
            /// 原因/建議-備註說明
            /// </summary>
            public string comment { get; set; } = "";

        }

        public class tbAppropriation
        {
            public String Form_no { get; set; } = "";
            public String ExamineNo { get; set; } = "";
            public String Appr_Type { get; set; } = "";
            public Int32 HandlingFee { get; set; } = 0;
            public Int32 PathFee { get; set; } = 0;
            public Int32 CustTotAmt { get; set; } = 0;
            public Int32 RemitFee { get; set; } = 0;
            public Int32 ActualAmt { get; set; } = 0;
            public Int32 PathAmt { get; set; } = 0;
            public String BankCode { get; set; } = "";
            public String BankName { get; set; } = "";
            public String BankID { get; set; } = "";
            public String AccountID { get; set; } = "";
            public String AccountName { get; set; } = "";
            public String Transfer_date { get; set; } = "";
            public String Add_User { get; set; } = "";
            public String Add_date { get; set; } = "";
            public String Upd_User { get; set; } = "";
            public String Upd_date { get; set; } = "";



        }


    }
}
