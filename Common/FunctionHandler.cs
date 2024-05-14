using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;

using System.Data.SqlClient;
using Business_Logic.Common;
using Business_Logic.Entity;
using System.Reflection;
using Business_Logic.Model;
using System.Collections;
using NPOI.SS.UserModel;
using System.IO;
using NPOI.HSSF.UserModel;
using NPOI.XSSF.UserModel;
using Newtonsoft.Json;
using System.Net.Mail;
using System.Configuration;
using System.IO.Compression;
using System.Text.RegularExpressions;
using System.Net;
using System.Net.Sockets;
using System.Web;
using System.Web.UI;
using System.Web.UI.HtmlControls;
using System.Web.UI.WebControls;
using System.Net.Http;
using System.Runtime.Remoting.Contexts;
using System.Net.Http.Headers;

namespace Business_Logic
{
    public class FunctionHandler
    {
        public DataComponent g_DC = new DataComponent();

        public Business_Logic.Model.BL_System g_BL_System = new BL_System();


        // public string CheckJsonType<T>(string propNa, T p_JsonVal)
        public void GetCtrlSetVal(string Variables, Control p_Control)
        {
            HtmlInputText RefText = new HtmlInputText();
            HtmlInputHidden RefHid = new HtmlInputHidden();
            TextBox RefTextBox=new TextBox();

            HtmlGenericControl RefCtrl = new HtmlGenericControl();

            Label RefLabel = new Label();
            HtmlImage RefImg = new HtmlImage();
            if ((!string.IsNullOrEmpty(Variables))&&(p_Control != null))
            {
                if (p_Control.GetType() == RefText.GetType())
                {
                    ((HtmlInputText)p_Control).Value = Variables;
                }
                if (p_Control.GetType() == RefHid.GetType())
                {
                    ((HtmlInputHidden)p_Control).Value = Variables;
                }
                if (p_Control.GetType() == RefTextBox.GetType())
                {
                    ((TextBox)p_Control).Text = Variables;
                }
                if (p_Control.GetType() == RefCtrl.GetType())
                {
                    ((HtmlGenericControl)p_Control).InnerText = Variables;
                }
                if (p_Control.GetType() == RefLabel.GetType())
                {
                    ((Label)p_Control).Text = Variables;
                }
                if (p_Control.GetType() == RefImg.GetType())
                {
                    ((HtmlImage)p_Control).Src = Variables;
                }
            }
        }


        public void Send_LINE(string MSG)
        {
            /*傳送LINE訊息*/
            System.Net.ServicePointManager.ServerCertificateValidationCallback = delegate { return true; };
            System.Net.ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12 | SecurityProtocolType.Ssl3 | SecurityProtocolType.Tls;

            HttpClient httpClient = new HttpClient();
            httpClient.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/x-www-form-urlencoded"));
            httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", "dHHOEdYKWAonV6bSmcdtStXuARKgcq968N5czjRIWdX");
            var content = new Dictionary<string, string>();
            content.Add("message", MSG);
            httpClient.PostAsync("https://notify-api.line.me/api/notify", new FormUrlEncodedContent(content));
        }

        public Boolean isMobileServer()
        {
            Boolean m_isMobileServer = false;
            try
            {
                var host = Dns.GetHostEntry(Dns.GetHostName());
                foreach (var ip in host.AddressList)
                {
                    if (ip.AddressFamily == AddressFamily.InterNetwork)
                    {
                        m_isMobileServer = ip.ToString() == ConfigurationManager.AppSettings["MobileIP"].ToString();
                    }
                }
            }
            catch
            {

            }
            return m_isMobileServer;
        }


        public Boolean GetDeviceIsMobile(string p_HTTP_USER_AGENT)
        {
            Boolean m_IsMobile = false;
            try
            {

                Regex b = new Regex(@"(android|bb\d+|meego).+mobile|avantgo|bada\/|blackberry|blazer|compal|elaine|fennec|hiptop|iemobile|ip(hone|od)|iris|kindle|lge |maemo|midp|mmp|mobile.+firefox|netfront|opera m(ob|in)i|palm( os)?|phone|p(ixi|re)\/|plucker|pocket|psp|series(4|6)0|symbian|treo|up\.(browser|link)|vodafone|wap|windows ce|xda|xiino", RegexOptions.IgnoreCase | RegexOptions.Multiline);
                Regex v = new Regex(@"1207|6310|6590|3gso|4thp|50[1-6]i|770s|802s|a wa|abac|ac(er|oo|s\-)|ai(ko|rn)|al(av|ca|co)|amoi|an(ex|ny|yw)|aptu|ar(ch|go)|as(te|us)|attw|au(di|\-m|r |s )|avan|be(ck|ll|nq)|bi(lb|rd)|bl(ac|az)|br(e|v)w|bumb|bw\-(n|u)|c55\/|capi|ccwa|cdm\-|cell|chtm|cldc|cmd\-|co(mp|nd)|craw|da(it|ll|ng)|dbte|dc\-s|devi|dica|dmob|do(c|p)o|ds(12|\-d)|el(49|ai)|em(l2|ul)|er(ic|k0)|esl8|ez([4-7]0|os|wa|ze)|fetc|fly(\-|_)|g1 u|g560|gene|gf\-5|g\-mo|go(\.w|od)|gr(ad|un)|haie|hcit|hd\-(m|p|t)|hei\-|hi(pt|ta)|hp( i|ip)|hs\-c|ht(c(\-| |_|a|g|p|s|t)|tp)|hu(aw|tc)|i\-(20|go|ma)|i230|iac( |\-|\/)|ibro|idea|ig01|ikom|im1k|inno|ipaq|iris|ja(t|v)a|jbro|jemu|jigs|kddi|keji|kgt( |\/)|klon|kpt |kwc\-|kyo(c|k)|le(no|xi)|lg( g|\/(k|l|u)|50|54|\-[a-w])|libw|lynx|m1\-w|m3ga|m50\/|ma(te|ui|xo)|mc(01|21|ca)|m\-cr|me(rc|ri)|mi(o8|oa|ts)|mmef|mo(01|02|bi|de|do|t(\-| |o|v)|zz)|mt(50|p1|v )|mwbp|mywa|n10[0-2]|n20[2-3]|n30(0|2)|n50(0|2|5)|n7(0(0|1)|10)|ne((c|m)\-|on|tf|wf|wg|wt)|nok(6|i)|nzph|o2im|op(ti|wv)|oran|owg1|p800|pan(a|d|t)|pdxg|pg(13|\-([1-8]|c))|phil|pire|pl(ay|uc)|pn\-2|po(ck|rt|se)|prox|psio|pt\-g|qa\-a|qc(07|12|21|32|60|\-[2-7]|i\-)|qtek|r380|r600|raks|rim9|ro(ve|zo)|s55\/|sa(ge|ma|mm|ms|ny|va)|sc(01|h\-|oo|p\-)|sdk\/|se(c(\-|0|1)|47|mc|nd|ri)|sgh\-|shar|sie(\-|m)|sk\-0|sl(45|id)|sm(al|ar|b3|it|t5)|so(ft|ny)|sp(01|h\-|v\-|v )|sy(01|mb)|t2(18|50)|t6(00|10|18)|ta(gt|lk)|tcl\-|tdg\-|tel(i|m)|tim\-|t\-mo|to(pl|sh)|ts(70|m\-|m3|m5)|tx\-9|up(\.b|g1|si)|utst|v400|v750|veri|vi(rg|te)|vk(40|5[0-3]|\-v)|vm40|voda|vulc|vx(52|53|60|61|70|80|81|83|85|98)|w3c(\-| )|webc|whit|wi(g |nc|nw)|wmlb|wonu|x700|yas\-|your|zeto|zte\-", RegexOptions.IgnoreCase | RegexOptions.Multiline);

                if ((b.IsMatch(p_HTTP_USER_AGENT) || v.IsMatch(p_HTTP_USER_AGENT.Substring(0, 4))))
                {
                    m_IsMobile = true;
                }
            }
            catch(Exception ex)
            { 
            }
            return m_IsMobile;
        }

        public Dictionary<string, object> DataSetToJson(DataSet p_DataSet)
        {
            Dictionary<string, object> m_Dictoa = new Dictionary<string, object>();
            try
            {
                foreach (DataTable dt in p_DataSet.Tables)
                {

                    m_Dictoa.Add(dt.TableName, RowsToDictionary(dt));
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            return m_Dictoa;
        }

        public Dictionary<string, object> DataTableToJson(DataTable p_DataTable)
        {
            Dictionary<string, object> m_Dictoa = new Dictionary<string, object>();
            try
            {
                m_Dictoa.Add(p_DataTable.TableName, RowsToDictionary(p_DataTable));
            }
            catch (Exception ex)
            {
                throw ex;
            }
            return m_Dictoa;
        }

        public string GetCityJSON()
        {
            // 讀取 JSON 檔案的內容
            string path = AppDomain.CurrentDomain.BaseDirectory;
            string m_Json = System.IO.File.ReadAllText(path+ @"\resource\AllRoad.json");

            return m_Json;  
        }

        public Dictionary<string, object> DataTableToDic(DataTable p_DataTable)
        {
            Dictionary<string, object> m_Dic = new Dictionary<string, object>();
            try
            {
                foreach (DataRow dr in p_DataTable.Rows)
                {
                    for (int i = 0; i < p_DataTable.Columns.Count; i++)
                    {
                        m_Dic.Add(p_DataTable.Columns[i].ColumnName, dr[i]);
                    }

                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            return m_Dic;
        }

        public T DataRowConvertEntity<T>(DataRow p_dr) where T : new()
        {

         var m_STypet = new T();
         try
         {
             if (p_dr != null)
             {
                 PropertyInfo[] properties = m_STypet.GetType().GetProperties();

                 foreach (PropertyInfo property in properties)
                 {
                     if (p_dr.Table.Columns.Contains(property.Name))
                     {
                         // Find which property type (int, string, double? etc) the CURRENT property is...
                         Type tPropertyType = m_STypet.GetType().GetProperty(property.Name).PropertyType;
                         // Fix nullables...
                         Type newT = Nullable.GetUnderlyingType(tPropertyType) ?? tPropertyType;
                         // ...and change the type
                         object newA = Convert.ChangeType(p_dr[property.Name], newT);
                         m_STypet.GetType().GetProperty(property.Name).SetValue(m_STypet, newA, null);
                     }
                    
                 }
             }
         
         }
         catch (Exception ex)
         {
             throw ex;
         }
         return m_STypet;
        }

        public List<Dictionary<string, object>> RowsToDictionary(DataTable p_dt)
        {
            List<Dictionary<string, object>> m_objs = new List<Dictionary<string, object>>();
            try
            {
                foreach (DataRow dr in p_dt.Rows)
                {
                    Dictionary<string, object> drow = new Dictionary<string, object>();
                    for (int i = 0; i < p_dt.Columns.Count; i++)
                    {
                        drow.Add(p_dt.Columns[i].ColumnName, dr[i]);
                    }
                    m_objs.Add(drow);
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            return m_objs;
        }
   
        public Type GetBLType(string m_TypeName)
        {
            try
            {
                return Type.GetType("Business_Logic.Model." + m_TypeName);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public T ConvertEntity<T>(IDictionary<string, object> dict, ref Dictionary<string, Object> p_objBetween) where T : new()
        {
            var m_STypet = new T();
            try
            {
                if (dict != null)
                {
                    PropertyInfo[] properties = m_STypet.GetType().GetProperties();

                    foreach (PropertyInfo property in properties)
                    {
                        if (!dict.Any(x => x.Key.Equals(property.Name, StringComparison.InvariantCultureIgnoreCase)))
                        {
                            try
                            {
                                p_objBetween[property.Name + "_S"] = dict[property.Name + "_S"].ToString();
                                p_objBetween[property.Name + "_E"] = dict[property.Name + "_E"].ToString();

                            }
                            catch
                            { }
                            continue;
                        }
                        KeyValuePair<string, object> item = dict.First(x => x.Key.Equals(property.Name, StringComparison.InvariantCultureIgnoreCase));
                        // Find which property type (int, string, double? etc) the CURRENT property is...
                        Type tPropertyType = m_STypet.GetType().GetProperty(property.Name).PropertyType;
                        // Fix nullables...
                        Type newT = Nullable.GetUnderlyingType(tPropertyType) ?? tPropertyType;
                        // ...and change the type
                        object newA = Convert.ChangeType(item.Value, newT);
                        m_STypet.GetType().GetProperty(property.Name).SetValue(m_STypet, newA, null);
                    }
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            return m_STypet;

        }


        public Dictionary<string, object> CheckViewAuth(string p_ViewDetailID, DataTable p_dtViewAuth, DataTable p_dtAuthID, DataTable p_DataTable, string p_AuthType)
        {
            Dictionary<string, object> m_Result = new Dictionary<string, object>();
            string m_ReadOnly = "";
            bool m_isVisible = false;
            string m_Style = "";
          
            if (p_dtViewAuth.Rows.Count != 0)
            {
                DataRow[] m_Rows = p_dtViewAuth.Select("ViewDetailID='" + p_ViewDetailID + "'");
                try
                {
                    if (m_Rows.Length != 0)
                    {
                        //權限管理
                        foreach (DataRow AuthDr in m_Rows)
                        {
                            bool m_isAuthRule = true;
                            if (p_AuthType != "Initial")
                            {
                                m_isAuthRule = p_DataTable.Rows[0][AuthDr["AuthRule"].ToString()].ToString() == AuthDr["AuthRuleValue"].ToString();
                            }

                            if (m_isAuthRule)
                            {
                                string m_SiteFormAuthID = AuthDr["AuthID"].ToString();
                                if (m_SiteFormAuthID == "ALL")
                                {
                                    if (p_AuthType == "Initial")
                                    {
                                        if (AuthDr["Visible"].ToString() == "Y" || AuthDr["Visible"].ToString() == "")
                                        {
                                            m_isVisible = true;
                                        }
                                        if (AuthDr["Edit"].ToString() == "N")
                                        {
                                            m_ReadOnly = " ReadOnly='true'  disabled='disabled' ";
                                        }
                                        else
                                        {
                                            m_ReadOnly = "  ";

                                        }
                                        m_Style = AuthDr["Style"].ToString();
                                    }
                                    else
                                    {
                                        if (p_DataTable.Rows.Count != 0)
                                        {
                                            if (m_isAuthRule)
                                            {
                                                if (AuthDr["Visible"].ToString() == "Y" || AuthDr["Visible"].ToString() == "")
                                                {
                                                    m_isVisible = true;
                                                }
                                                else
                                                {
                                                    m_isVisible = false;
                                                }
                                                if (AuthDr["Edit"].ToString() == "N")
                                                {
                                                    m_ReadOnly = " ReadOnly='true'  disabled='disabled' ";
                                                }
                                                else
                                                {
                                                    m_ReadOnly = "  ";

                                                }
                                                m_Style = AuthDr["Style"].ToString();
                                            }
                                            else
                                            {
                                                if (AuthDr["Visible"].ToString() == "N" || AuthDr["Visible"].ToString() == "")
                                                {
                                                    m_isVisible = true;
                                                }
                                                else
                                                {
                                                    m_isVisible = false;
                                                }
                                                if (AuthDr["Edit"].ToString() == "Y")
                                                {
                                                    m_ReadOnly = " ReadOnly='true'  disabled='disabled' ";
                                                }
                                                else
                                                {
                                                    m_ReadOnly = "  ";

                                                }
                                                m_Style = "";

                                            }
                                        }
                                    }
                                }
                                else
                                {
                                    DataRow[] m_UserAuthDrs = p_dtAuthID.Select("AuthID='" + m_SiteFormAuthID + "'");
                                    if (m_UserAuthDrs.Length != 0)
                                    {//符合權限
                                        if (p_AuthType == "Initial")
                                        {
                                            if (AuthDr["Visible"].ToString() == "Y" || AuthDr["Visible"].ToString() == "")
                                            {
                                                m_isVisible = true;
                                            }
                                            else
                                            {
                                                m_isVisible = false;
                                            }

                                            if (AuthDr["Edit"].ToString() == "N")
                                            {
                                                m_ReadOnly = " ReadOnly='true'  disabled='disabled' ";
                                            }
                                            else
                                            {
                                                m_ReadOnly = "  ";
                                                break;
                                            }
                                            m_Style = AuthDr["Style"].ToString();
                                            break;
                                        }
                                        else
                                        {
                                            if (p_DataTable.Rows.Count != 0)
                                            {
                                                if (m_isAuthRule)
                                                {
                                                    if (AuthDr["Visible"].ToString() == "Y" || AuthDr["Visible"].ToString() == "")
                                                    {
                                                        m_isVisible = true;
                                                    }
                                                    else
                                                    {
                                                        m_isVisible = false;
                                                    }
                                                    if (AuthDr["Edit"].ToString() == "N")
                                                    {
                                                        m_ReadOnly = " ReadOnly='true'  disabled='disabled' ";
                                                    }
                                                    else
                                                    {
                                                        m_ReadOnly = "  ";
                                                        break;
                                                    }
                                                    m_Style = AuthDr["Style"].ToString();
                                                }
                                                else
                                                {
                                                    if (AuthDr["Visible"].ToString() == "N" || AuthDr["Visible"].ToString() == "")
                                                    {
                                                        m_isVisible = true;
                                                    }
                                                    else
                                                    {
                                                        m_isVisible = false;
                                                    }
                                                    if (AuthDr["Edit"].ToString() == "Y")
                                                    {
                                                        m_ReadOnly = " ReadOnly='true'  disabled='disabled' ";
                                                    }
                                                    else
                                                    {
                                                        m_ReadOnly = "  ";

                                                    }
                                                    m_Style = "";
                                                }
                                            }
                                        }
                                        break;
                                    }
                                    else
                                    {//不符合權限
                                       
                                            if (AuthDr["Visible"].ToString() == "N" || AuthDr["Visible"].ToString() == "")
                                            {
                                                m_isVisible = true;
                                            }
                                            else
                                            {
                                                m_isVisible = false;
                                            }
                                            if (AuthDr["Edit"].ToString() == "Y")
                                            {
                                                m_ReadOnly = " ReadOnly='true'  disabled='disabled' ";
                                            }
                                        

                                    }
                                }
                            }
                        }
                    }
                    else
                    {
                        m_isVisible = true;
                        m_ReadOnly = "";
                        m_Style = "";
                    }
                }
                catch (Exception ex)
                {
                    m_Result["Error"] = ex.Message;
                }
            }
            else
            {
                m_isVisible = true;
                m_ReadOnly = "";
                m_Style = "";
            }
            m_Result["isVisible"] = m_isVisible;
            m_Result["ReadOnly"] = m_ReadOnly;
            m_Result["Style"] = m_Style;
            return m_Result;
        }

        public string GeneratorForm(SysEntity.Employee p_Employee, string p_SiteFormID, string p_SectionID, string p_Language, string p_ActionMode = "", DataTable p_DataTable = null, DataTable p_FormTable = null, string p_PreviewKep = "", Dictionary<string, bool> p_SiteFromStatus = null,bool p_IsMobile=false)
        {
            string m_ControlHTML = "";
            try
            {
                string m_TabID = "BaseCtrl";
                bool m_isSiteFormID = true;

                DataTable m_ViewDetail = new DataTable();

                switch (p_SiteFormID)
                {
                    case "Detail":
                        m_ViewDetail = p_FormTable.Copy();
                        m_TabID = "DetailFormCtrl";
                        m_isSiteFormID = false;
                        break;
                    case "QueryForm":
                        m_ViewDetail = p_FormTable.Copy();
                        m_TabID = "QueryFormCtrl";
                        m_isSiteFormID = false;
                        break;
                    default:
                        m_ViewDetail = g_BL_System.GetViewDetail(p_Employee, p_SiteFormID, p_SectionID, "", p_PreviewKep);
                        break;
                }

            
                string m_SectionID = p_SectionID;
                m_ControlHTML = "<table id='" + m_TabID + "' class='' style='width:95%;font-size:small;border-color: black; border-collapse: collapse;' >";
                string m_SectionHtml = "";
                string m_Control = "";
                //bool m_isDetail = false;
                Int32 m_ItemIndex = 1;
                Int32 m_TDIndex = 1;
                Int32 m_ColumnCount = 0;
                Int32 m_AllItemCount = m_ViewDetail.Rows.Count;
                DataTable m_dtViewAuths = new DataTable();
                string m_AuthType = "";
                if (p_ActionMode != "" && p_ActionMode != "Add" || p_SectionID=="divInitButton")
                {
                    try
                    {
                        if (m_isSiteFormID)
                        {
                            if (p_SectionID == "divInitButton")
                            {
                                m_AuthType = "Initial";
                                m_dtViewAuths = g_BL_System.GetViewAuthByAuthKey(p_Employee, p_SiteFormID, "Initial");
                            }
                            else
                            {
                                m_AuthType = "Maintain";
                                m_dtViewAuths = g_BL_System.GetViewAuthByAuthKey(p_Employee, p_SiteFormID, "Maintain");
                            }
                        }
                    }
                    catch
                    { }
                }
                DataTable m_dtAuthID = new DataTable();
                if (m_dtViewAuths.Rows.Count != 0)
                {
                    m_dtAuthID = g_BL_System.GetAuthIDByWorkID(p_Employee, p_Employee.WorkID);
                }
                string m_SectionType = "";
                string m_LanguageValue = "";
                foreach (DataRow m_dr in m_ViewDetail.Rows)
                {
                    bool m_isLastItem = m_AllItemCount == m_ItemIndex;
                    string m_ControlType = m_dr["ControlType"].ToString();
                    m_SectionType = m_dr["SectionType"].ToString();
                    string m_FieldID = m_dr["FieldID"].ToString();
                    string m_QueryID = m_dr["QueryID"].ToString();
                    string m_ControlRef = m_dr["ControlRef"].ToString();
                    string m_ValidID = m_dr["ValidID"].ToString();
                    string m_Height = m_dr["Height"].ToString();
                    string m_Width = m_dr["Width"].ToString();
                    string m_EventID = m_dr["EventID"].ToString();
                    string m_EventAction = m_dr["EventAction"].ToString();
                    string m_EventActionType = m_dr["EventActionType"].ToString();
                    string m_ValidSource = m_dr["ValidSource"].ToString();
                    string m_ValidParam = m_dr["ValidParam"].ToString();
                    string m_Valid = "";
                    string m_MaxLength = m_dr["MaxLength"].ToString();
                    string m_ModifyKey = m_dr["ModifyKey"].ToString();
                    string m_EventFollow = m_dr["EventFollow"].ToString();
                    string m_QueryMethod = m_dr["QueryMethod"].ToString();
                    string m_ReadOnly = "";
                    string m_ViewDetailID =  m_dr["ViewDetailID"].ToString();
                    string m_DefaultValue = m_dr["DefaultValue"].ToString();
                    string m_AuthStyle = "";
                    if (p_IsMobile && m_SectionType.IndexOf("General")==0)
                    {
                        m_SectionType = "General";
                    }
                    if (m_ModifyKey == "1" || p_ActionMode != "Add")
                    {
                        if (m_ModifyKey == "1" && p_ActionMode != "Add")
                        {
                            m_ReadOnly = " ReadOnly='true'  disabled='disabled' ";
                        }
                        if (p_ActionMode != "" && p_ActionMode != "Add" || p_SectionID == "divInitButton")
                        {
                            //權限管理
                            if (m_dtViewAuths.Rows.Count != 0)
                            {
                                Int32 m_AuthCount = m_dtViewAuths.Select("ViewDetailID='" + m_ViewDetailID + "'").Length;
                                if (m_AuthCount != 0)
                                {
                                    Dictionary<string, object> m_dicAuthResult = CheckViewAuth(m_ViewDetailID, m_dtViewAuths, m_dtAuthID, p_DataTable, m_AuthType);
                                    if (!(m_dicAuthResult.ContainsKey("Error")))
                                    {
                                        if (!(bool)m_dicAuthResult["isVisible"])
                                        {
                                            continue;
                                        }
                                        if (m_dicAuthResult.ContainsKey("ReadOnly"))
                                        {
                                            m_ReadOnly = m_dicAuthResult["ReadOnly"].ToString();
                                        }
                                        if (m_dicAuthResult.ContainsKey("Style"))
                                        {
                                            m_AuthStyle = m_dicAuthResult["Style"].ToString();
                                        }
                                    }
                                }
                            }
                        }
                    }
                    if ( p_ActionMode == "Add")
                    {
                        if (m_ModifyKey == "2")
                        {
                            m_ReadOnly = " ReadOnly='true'  disabled='disabled' ";
                        }
                    }
                    if (m_ValidSource != "")
                    {
                        m_Valid = " ValidType='" + m_ValidSource + "'  onfocusout='" + m_ValidSource + "(this)' ";
                    }
                    m_LanguageValue = GetMultiLanguage(m_FieldID, p_Language, p_Employee);


                   // if (m_isDetail == false)
                    //{
                    m_Control = GeneratorControl(p_Employee, p_ActionMode, p_SiteFormID, m_ControlType, m_SectionType, m_FieldID, m_Valid, m_EventAction, p_Language, m_LanguageValue, m_MaxLength, m_Width, m_Height, m_QueryID, m_EventID, m_ControlRef, m_EventActionType, m_SectionID, m_ReadOnly, m_EventFollow, m_QueryMethod, m_ModifyKey, p_DataTable, "", m_DefaultValue, m_AuthStyle, p_IsMobile);
                   // }
                    if (m_SectionID == "divInitButton")
                    {
                        m_SectionHtml += " " + m_Control + " ";
                    }
                    else
                    {
                        if (m_SectionType != "Detail")
                        {
                         //   m_isDetail = false;
                            DataTable m_dtParam = GetParameter(m_SectionType, "SectionType", p_Employee);

                            foreach (DataRow drPar in m_dtParam.Rows)
                            {
                                m_ColumnCount = Convert.ToInt32(drPar["Value"].ToString());
                            }

                            //    m_SectionHtml = GeneratorTR(m_SectionHtml, m_SectionType, m_ControlType, m_Control, m_LanguageValue, m_FieldID, m_ColumnCount, m_Width, m_Height, m_isLastItem, m_ItemIndex, m_TDIndex, "", m_AuthStyle, m_ValidSource);
                            //}
                            //else
                            //{
                            //    if (m_isDetail == false)
                            //    {
                            //        m_isDetail = true;
                            //        m_SectionHtml = GeneratorTR(m_SectionHtml, m_SectionType, m_ControlType, m_Control, m_LanguageValue, m_FieldID, m_ColumnCount, m_Width, m_Height, m_isLastItem, m_ItemIndex, m_TDIndex, "", m_AuthStyle, m_ValidSource);
                            //    }
                        }
                        m_SectionHtml = GeneratorTR(m_SectionHtml, m_SectionType, m_ControlType, m_Control, m_LanguageValue, m_FieldID, m_ColumnCount, m_Width, m_Height, m_isLastItem, m_ItemIndex, m_TDIndex, "", m_AuthStyle, m_ValidSource);
                        

                    }
                    m_ItemIndex++;
                    m_TDIndex++;
                    if (m_Width == "100%")
                    {
                        m_TDIndex = 1;
                    }
                }

                if (m_SectionID == "divInitButton")
                {
                    m_SectionHtml = "<tr><td  style='text-align:left;' >" + m_SectionHtml + "</td></tr>";
                }

                //簽核狀態
                if (m_SectionID == "divQueryArea")
                {
                    if (p_SiteFromStatus != null)
                    {
                        foreach (KeyValuePair<string, bool> item in p_SiteFromStatus)
                        {
                            string m_FieldID = item.Key;
                            if (item.Value)
                            {

                                DataTable m_dtStatus = g_BL_System.GetSiteFromStatusBySiteFromID(p_Employee, p_SiteFormID, m_FieldID);
                                if (m_dtStatus.Rows.Count != 0)
                                {
                                    string m_StatusTitle = m_dtStatus.Rows[0]["StatusTitle"].ToString();
                                    m_LanguageValue = GetMultiLanguage(m_StatusTitle, p_Language, p_Employee);

                                    m_Control = "";
                                    string m_LabelHtml = "<label id='lbl{0}' sel='' class='SiteFormStatus' StatusFieldID='" + m_FieldID + "' ondblclick='return false'  onclick='ActionModel(\"{1}\",\"" + m_FieldID + "\",\"{2}\",\"\",\"\",\"{3}\");SiteFormStatusOnclick(this);InitSearchInfo(\"{4}\",\"" + m_FieldID + "\",\"{5}\",\"{6}\");' >{7}</label>　　";
                                    foreach (DataRow dr in m_dtStatus.Rows)
                                    {
                                        string m_Text = GetMultiLanguage(dr["StatusKey"].ToString(), p_Language, p_Employee);
                                        m_Control += string.Format(m_LabelHtml, dr["StatusKey"].ToString(), dr["EventID"].ToString(), dr["EventActionType"].ToString(), dr["ConditionRef"].ToString() + ";" + dr["ConditionType"].ToString() + ";" + dr["StatusCondition"].ToString(), dr["EventID"].ToString(), dr["EventActionType"].ToString(), dr["ConditionRef"].ToString() + ";" + dr["ConditionType"].ToString() + ";" + dr["StatusCondition"].ToString(), m_Text);
                                    }
                                    m_SectionHtml = GeneratorTR(m_SectionHtml, m_SectionType, "SiteFormStatus", m_Control, m_LanguageValue, m_StatusTitle, m_ColumnCount, "", "", true, m_ItemIndex, m_TDIndex);


                                    
                                
                                }
                            }

                        }
                    }
                    
                   
                }
                m_ControlHTML += m_SectionHtml;
                m_ControlHTML += "</table>";
            }
            catch (Exception ex)
            {
                throw ex;
            }
            return m_ControlHTML;
        }

        public string GetValueFormat(string p_ControlType, string p_DefaultValue,SysEntity.Employee p_Employee)
        {
            string m_Value = "";
            if (p_DefaultValue == "TimeKey")
            {
                p_DefaultValue = DateTime.Now.ToString("yyyyMMddHHmmssfff");
            }
            if (p_DefaultValue == "SysDate")
            {
                p_DefaultValue = DateTime.Now.ToString("yyyyMMdd");
            }
            

            switch (p_ControlType)
            {
                case "Display":
                case "Select":
                case "CheckBox":
                case "MultiText":
                    m_Value = p_DefaultValue;
                    break;
                case "DateTime":
                case "Date":
                    if (p_DefaultValue != "")
                    {
                        if (p_DefaultValue.IndexOf("/") == -1)
                        {
                            m_Value = " value='" + p_DefaultValue.Insert(4, @"/").Insert(7, @"/") + "' ";
                        }
                        else
                        {
                            m_Value = " value='" + p_DefaultValue + "' ";
                        }
                    }
                    else
                    {
                        m_Value = " value='" + p_DefaultValue + "' ";
                    }
                    
                    break;
                default:
                    m_Value = " value='" + p_DefaultValue + "' ";

                    break;

            }
            return m_Value;
        }

        public string chkDefaultValue( SysEntity.Employee p_Employee ,string p_DefaultValue)
        {
            string m_DefaultValue = "";
            if (p_DefaultValue == "Employee.WorkID")
            {
                m_DefaultValue = "value='"+ p_Employee.WorkID+"'";
            }

            if (p_DefaultValue == "Employee.CompanyCode")
            {
                m_DefaultValue = "value='" + p_Employee.CompanyCode + "'";
            }


            return m_DefaultValue;  


        }
        public string GeneratorControl(SysEntity.Employee p_Employee, string p_ActionMode, string p_SiteFormID, string p_ControlType, string p_SectionType, string p_FieldID, string p_Valid, string p_EventAction, string p_Language, string p_LanguageValue, string p_MaxLength, string p_Width, string p_Height, string p_QueryID, string p_EventID, string p_ControlRef, string p_EventActionType, string p_SectionID, string p_ReadOnly, string p_EventFollow, string p_QueryMethod, string p_ModifyKey, DataTable p_DataTable = null, string p_DtlControlID = "", string p_DefaultValue = "", string p_AuthStyle = "", bool p_IsMobile = false)
        {
            string m_Control = "";
            string m_style = "style='{0}' ";
            string m_MaxLength = "";
            string m_ControlRef = "";
            string m_DataValue = "";
            string m_OnClick = "";
            string m_DefaultValue = p_DefaultValue;
            //use by Select MultiSelect 
            string m_Selected = "";
            string m_Options = "";
            string m_selValue = "";
            string m_selText = "";
            string m_OnChange = "";
            string m_OnlyValue = "";
            Dictionary<string, Object> m_DisDefVal = new  Dictionary<string, Object>();
            try
            {
                string m_chkDefaultValue = "";
                if (p_ActionMode == "Add")
                {
                    if (p_DefaultValue != "")
                    {
                        m_chkDefaultValue = chkDefaultValue(p_Employee, p_DefaultValue);
                        if (m_chkDefaultValue == "")
                        {
                            object[] m_Params = new object[1];
                            m_Params[0] = p_Employee;
                            SysEntity.TransResult m_TransResultDV = GetDefaultValue(p_DefaultValue, m_Params, p_Employee);
                            if (m_TransResultDV.isSuccess)
                            {
                                try
                                {
                                    if (m_TransResultDV.ResultEntity.GetType() == m_DisDefVal.GetType())
                                    {
                                        m_DisDefVal = (Dictionary<string, Object>)m_TransResultDV.ResultEntity;
                                    }
                                    else
                                    {
                                        m_DefaultValue = (string)m_TransResultDV.ResultEntity;
                                    }
                                }
                                catch
                                {
                                    throw new Exception("FieldID:" + p_FieldID + " Set DefaultValue is Error!! check  ViewDetail ");
                                }
                            }
                            m_DataValue = GetValueFormat(p_ControlType, m_DefaultValue, p_Employee);
                        }
                        else 
                        {
                            m_DataValue = m_chkDefaultValue;
                        }
                    }
                }
                else
                {
                    if (p_DefaultValue != "" && p_ControlType == "Query")
                    {
                        try
                        {
                            m_chkDefaultValue = chkDefaultValue(p_Employee, p_DefaultValue);
                            if (m_chkDefaultValue == "")
                            {
                                object[] m_Params = new object[1];
                                SysEntity.Employee m_ElseEmp = new SysEntity.Employee();
                                m_ElseEmp.WorkID = p_DataTable.Rows[0][p_FieldID].ToString();
                                m_Params[0] = m_ElseEmp;
                                SysEntity.TransResult m_TransResultDV = GetDefaultValue(p_DefaultValue, m_Params, p_Employee);
                                if (m_TransResultDV.isSuccess)
                                {
                                    try
                                    {
                                        if (m_TransResultDV.ResultEntity.GetType() == m_DisDefVal.GetType())
                                        {
                                            m_DisDefVal = (Dictionary<string, Object>)m_TransResultDV.ResultEntity;
                                        }
                                        else
                                        {
                                            m_DefaultValue = (string)m_TransResultDV.ResultEntity;
                                        }
                                    }
                                    catch
                                    {
                                        throw new Exception("FieldID:" + p_FieldID + " Set DefaultValue is Error!! check  ViewDetail ");
                                    }
                                }
                                m_DataValue = GetValueFormat(p_ControlType, m_DefaultValue, p_Employee);
                            }
                            else
                            {
                                m_DataValue = m_chkDefaultValue;
                            }


                            
                        }
                        catch
                        { 
                        
                        }
                       

                    }
                }
                if (p_DataTable != null)
                {
                    DataColumnCollection columns = p_DataTable.Columns;

                    foreach (DataRow dr in p_DataTable.Rows)
                    {
                        try
                        {
                            if (columns.Contains(p_FieldID))
                            {
                                m_OnlyValue = dr[p_FieldID].ToString();
                                m_DataValue = dr[p_FieldID].ToString();
                            }
                        }
                        catch
                        {
                            //throw new Exception("FieldID:" + p_FieldID + " Set Value is Error!!");
                        }
                        m_DataValue = GetValueFormat(p_ControlType, m_DataValue, p_Employee);
                    }
                }
                string m_ControlID = p_FieldID;
                if (p_DtlControlID != "")
                {
                    m_ControlID = p_DtlControlID;
                    m_style = "";
                }
                else
                {
                    if (p_Width == "")
                    {
                        m_style = string.Format(m_style, "width:90%;");
                    }
                    else
                    {
                        m_style = string.Format(m_style, p_Width);
                        if (p_Width == "100%")
                        {
                            m_style = string.Format(m_style, "width:90%;");
                        }
                        if (p_Width == "0")
                        {
                            m_style = string.Format(m_style, "display:none;");
                        }
                        if (p_Width == "display:none")
                        {
                            m_style = string.Format(m_style, "display:none;");
                        }
                    }
                }
                if (p_MaxLength != "")
                {
                    m_MaxLength = " maxlength='" + p_MaxLength + "' ";
                }
                if (p_AuthStyle != "")
                {
                    if (m_style != "")
                    {
                        m_style = m_style.Remove(m_style.Length-2) + p_AuthStyle + "'";
                    }
                }
                if (p_ControlRef != "")
                {
                    m_ControlRef = "ControlRef='" + p_ControlRef + "'";
                }
                string m_TextType = "";
                if (p_ControlType == "Date" || p_ControlType == "DateTime" || p_ControlType == "BetweenDate")
                {
                    m_TextType = " TextType='Date' ";
                    m_style = string.Format(m_style, "width:120px;");
                }
                string m_QueryMode = "Q";
                if (p_SectionID == "divMaintain")
                {
                    m_QueryMode = "";
                }
                if (p_SectionID == "divDetail")
                {
                    m_QueryMode = "Dtl";
                }
                string m_ModifyKey = "";
                if (p_ModifyKey == "1")
                {
                    m_ModifyKey = " ModifyKey='1' ";
                }
                if (p_SectionType != "Detail")
                {
                    switch (p_ControlType)
                    {
                        case "MultiSelect":
                            m_OnClick = " onclick='OpenQueryForm(\"" + m_QueryMode + "Qry" + m_ControlID + "\",\"" + p_ControlType + "\",\"" + p_QueryID + "\",\"" + p_ControlRef + "\",\"\",\"" + m_QueryMode + "\")'";
                           
                            if (p_DataTable != null)
                            {
                                object[] m_Params = new object[2];
                                m_Params[0] = p_Employee;
                                m_Params[1] = null;
                                DataTable m_dtSouce = GetQueryMethod(p_QueryMethod, m_Params, p_Employee);
                                if (p_ControlRef != "" && p_ControlRef.IndexOf(";")!=-1)
                                {
                                    m_selText = p_ControlRef.Split(';')[0].ToString().Split(',')[1];
                                    m_selValue = p_ControlRef.Split(';')[0].ToString().Split(',')[0];
                                }
                                foreach (DataRow EvRow in m_dtSouce.Rows)
                                {
                                    m_Selected = "";
                                    if (EvRow[m_selValue].ToString() == m_DataValue)
                                    {
                                        m_Selected = "selected";
                                    }
                                    m_Options += "<option value='" + EvRow[m_selValue].ToString() + "' " + m_Selected + ">" + EvRow[m_selText].ToString() + "</option>";
                                }
                            }
                            m_Control = "<table><tr><td class='qry'><SELECT class='MultiSelect' id='" + m_QueryMode + "Qry" + m_ControlID + "' " + m_DataValue + "  " + m_TextType + " " + p_ReadOnly + " " + p_Valid + " style='width:100%'  " + m_MaxLength + "  multiple='multiple'>" + m_Options + "</select></td><td><img name='" + m_QueryMode + "Qry" + m_ControlID + "' class='MultiSelect' " + m_OnClick + " src='Img/Query.png' alt='Query' height='20' width='20'></td></tr></table>";
                            break;
                        case "HtmlEditor":
                            m_Control = " <textarea class='HtmlEditor' id='HtmlEdi" + m_ControlID + "' rows='15' cols='50'></textarea>";
                            break;
                        case "DropUpload":
                            if (p_ReadOnly.Trim() == "")
                            {
                                string m_DropMsg = "";
                                string[] m_Msg = { "MSG00005" };
                                SysEntity.TransResult m_Lan = g_BL_System.GetListMultiLanguage(p_Employee, p_Language, m_Msg);
                                if (m_Lan.isSuccess)
                                {
                                    m_DropMsg = ((Dictionary<string, Object>)m_Lan.ResultEntity)["MSG00005"].ToString();
                                }
                                m_Control = "  <div id='Dfile" + m_ControlID + "' KeyID='" + p_ControlRef + "' FieldID='" + p_FieldID + "' Class='DropFile'><label id='lblMSG00005'>" + m_DropMsg + "</label>  </div>";
                            }
                            string m_Files = "";
                            if (p_DataTable != null)
                            {
                                if (p_DataTable.Rows.Count != 0)
                                {
                                    string m_FileKey = "";
                                    foreach (DataRow dr in p_DataTable.Rows)
                                    {
                                        m_FileKey = dr[p_ControlRef].ToString();
                                    }
                                    SysEntity.FileFolder m_FileFolder = new SysEntity.FileFolder();
                                    m_FileFolder.SiteFormID = p_SiteFormID;
                                    m_FileFolder.DocID = m_FileKey;
                                    m_FileFolder.FieldID = p_FieldID;
                                    SysEntity.TransResult m_TransResult = g_BL_System.GetFileFolder(p_Employee, m_FileFolder);
                                    if (m_TransResult.isSuccess)
                                    {
                                        foreach (DataRow dr in ((DataTable)m_TransResult.ResultEntity).Rows)
                                        {
                                            string m_DelImg = "";
                                            if (p_ReadOnly.Trim() == "ReadOnly='true'  disabled='disabled'")
                                            {
                                                m_DelImg = "";
                                            }
                                            else
                                            {
                                                m_DelImg = "<img src='Img/DelIco.png' alt='Del' onclick='DelFile(\"Dfile" + m_ControlID + "\",\"" + m_FileKey + "\",\"" + dr["FileKey"] + "\",\"" + p_SiteFormID + "\",\"" + p_FieldID + "\")'/>";
                                            }
                                            if (dr["FileType"].ToString().IndexOf("image/") != 0)
                                            {
                                                //m_Files += "<p class='DownLoad' onclick='DownloadDOC(\"" + dr["FileKey"] + "\",\"" + p_SiteFormID + "\")'>" + dr["FileName"].ToString() + "</p>";
                                                m_Files += "<p>" + m_DelImg + "<span class='DownLoad' onclick='ViewerDOC(\"" + dr["FileKey"] + "\",\"" + p_SiteFormID + "\")'>" + dr["FileName"].ToString() + "</span></p>";
                                            }
                                            else
                                            {
                                                m_Files += "<p>" + m_DelImg + "<span class='DownLoad' onclick='OpenImage(\"" + dr["FileKey"] + "\",\"" + p_SiteFormID + "\")'>" + dr["FileName"].ToString() + "</span></p>";
                                            }
                                        }
                                        m_Control += "<div id='Down" + m_ControlID + "'>" + m_Files + "</div>";
                                    }
                                }
                            }
                            break;

                        case "ExcelUpload":
                            m_Control = "  <input type='file'  id='file" + m_ControlID + "' />";
                            m_Control += " <input type='button'  onclick='UploadExcelGrid(\"file" + m_ControlID + "\",\"" + p_SiteFormID + "\",\"" + p_ControlRef + "\",\"" + p_FieldID + "\");' value='Upload Excel' />";
                            break;
                        case "FileUpload":
                            if (p_ReadOnly.Trim() == "")
                            {
                                m_Control = " <input class='gui-file' onchange='fileOnChange(this,$(\"#upl" + m_ControlID + "\")[0],\"file" + m_ControlID + "\",\"" + p_SiteFormID + "\",\"" + p_ControlRef + "\",\"" + p_FieldID + "\")' style='display:inline;' type='file' id='file" + m_ControlID + "' />&nbsp;&nbsp;";
                                m_Control += "<img class='hand' style='display:none'  id='upl" + m_ControlID + "'  src='Img/Upload.png' alt='Upload'  onclick='UploadFileAsync(this,\"file" + m_ControlID + "\",\"" + p_SiteFormID + "\",\"" + p_ControlRef + "\");' height='30' width='30'/> ";
                            }
                            string m_Files1 = "";
                            if (p_DataTable != null)
                            {
                                if (p_DataTable.Rows.Count != 0)
                                {
                                    string m_FileKey = "";
                                    foreach (DataRow dr in p_DataTable.Rows)
                                    {
                                        m_FileKey = dr[p_ControlRef].ToString();
                                    }
                                    SysEntity.FileFolder m_FileFolder = new SysEntity.FileFolder();
                                    m_FileFolder.SiteFormID = p_SiteFormID;
                                    m_FileFolder.DocID = m_FileKey;
                                    m_FileFolder.FieldID = p_FieldID;

                                    SysEntity.TransResult m_TransResult = g_BL_System.GetFileFolder(p_Employee, m_FileFolder);
                                    if (m_TransResult.isSuccess)
                                    {
                                        foreach (DataRow dr in ((DataTable)m_TransResult.ResultEntity).Rows)
                                        {
                                            string m_DelImg = "";
                                            if (p_ReadOnly.Trim() == "ReadOnly='true'  disabled='disabled'")
                                            {
                                                m_DelImg = "";
                                            }
                                            else
                                            {
                                                m_DelImg = "<img src='Img/DelIco.png' alt='Del' onclick='DelFile(\"upl" + m_ControlID + "\",\"" + m_FileKey + "\",\"" + dr["FileKey"] + "\",\"" + p_SiteFormID + "\",\"" + p_FieldID + "\")'/>";
                                            }
                                            if (dr["FileType"].ToString().IndexOf("image/") != 0)
                                            {
                                                //m_Files += "<p class='DownLoad' onclick='DownloadDOC(\"" + dr["FileKey"] + "\",\"" + p_SiteFormID + "\")'>" + dr["FileName"].ToString() + "</p>";
                                                m_Files1 += "<p>" + m_DelImg + "<span class='DownLoad' onclick='ViewerDOC(\"" + dr["FileKey"] + "\",\"" + p_SiteFormID + "\")'>" + dr["FileName"].ToString() + "</span></p>";
                                            }
                                            else
                                            {
                                                m_Files1 += "<p>" + m_DelImg + "<span class='DownLoad' onclick='OpenImage(\"" + dr["FileKey"] + "\",\"" + p_SiteFormID + "\")'>" + dr["FileName"].ToString() + "</span></p>";
                                            }
                                            //if (dr["FileType"].ToString().IndexOf("image/") != 0)
                                            //{
                                            //    //m_Files += "<p class='DownLoad' onclick='DownloadDOC(\"" + dr["FileKey"] + "\",\"" + p_SiteFormID + "\")'>" + dr["FileName"].ToString() + "</p>";
                                            //    m_Files1 += "<p class='DownLoad' onclick='ViewerDOC(\"" + dr["FileKey"] + "\",\"" + p_SiteFormID + "\")'>" + dr["FileName"].ToString() + "</p>";
                                            //}
                                            //else
                                            //{
                                            //    m_Files1 += "<p class='DownLoad' onclick='OpenImage(\"" + dr["FileKey"] + "\",\"" + p_SiteFormID + "\")'>" + dr["FileName"].ToString() + "</p>";
                                            
                                            //}
                                        }
                                        m_Control += "<div id='Down" + m_ControlID + "'>" + m_Files1 + "</div>";
                                    }
                                }
                            }
                            break;
                        case "BetweenText":
                        case "BetweenDate":
                            m_Control = "<table id='BetText" + m_ControlID + "'><tr ><td class='between'><input  " + m_ModifyKey + "  id='Qtxt" + m_ControlID + "_S'  " + p_ReadOnly + "  " + m_TextType + "  " + p_Valid + " " + m_style + "  " + m_MaxLength + "  type='text' /></td><td class='between'>～</td><td class='between'><input id='Qtxt" + m_ControlID + "_E'  " + p_ReadOnly + " " + m_TextType + "    " + p_Valid + " " + m_style + "  " + m_MaxLength + "  type='text' /></td></tr></table>";
                            break;
                        case "TitleGrid":
                        case "TitleTable":
                            string m_TableHtml = "";
                            string m_hidCtrls = "";
                            
                            if (p_QueryMethod != "" && p_DataTable != null)
                            {
                                //if (p_DataTable.Rows.Count != 0    ) //capser modify for 表單新增時需秀Grid 但不須秀 SignRecord,SysNote
                                if (p_DataTable.Rows.Count != 0 || (p_FieldID != "SignRecord" && p_FieldID != "SysNote"))
                                {
                                   
                                    DataTable m_dtParam = GetParameter(p_SectionType, "SectionType", p_Employee);
                                    Int32 m_ColumnCount = 0;
                                    foreach (DataRow drPar in m_dtParam.Rows)
                                    {
                                        m_ColumnCount = Convert.ToInt32(drPar["Value"].ToString());
                                    }
                                    m_TableHtml += "<div class='panel-heading text-center' style='max-width: 1400px; min-width: 800px;' > <span class='panel-title'><label id='lbl" + m_ControlID + "'>" + p_LanguageValue + "</label> </span> </div>";
                                    m_TableHtml += "<div class='panel-body'>";
                                   
                                    if (p_ControlType == "TitleTable")
                                    {
                                        string m_hidProp = "<input id='TitleHid{0}' value='{1}' type='hidden' />";
                                        string[] m_hidProps = { "" };//TitleTable 隱藏控制項
                                        string m_QueryColumn = "";

                                        Int32 m_Ref = 1;
                                        foreach (string Prop in p_ControlRef.Split(';'))
                                        {
                                            if (m_Ref == 1)//TitleTable 隱藏控制項
                                            {
                                                m_hidProps = Prop.Split(',');
                                            }
                                            else if (m_Ref == 2)//TitleTable 查詢視窗定義
                                            {
                                                m_QueryColumn = Prop;
                                            }
                                            m_Ref++;
                                        }
                                        foreach (string Prop in p_ControlRef.Split(';'))
                                        {
                                            m_hidProps = Prop.Split(',');
                                        }
                                        object[] m_Params = new object[2];
                                        m_Params[0] = p_Employee;
                                        m_Params[1] = DataTableToDic(p_DataTable);
                                        DataTable m_dtSouce = GetQueryMethod(p_QueryMethod, m_Params, p_Employee);
                                        m_TableHtml += "<div class='TitleTableDivBody'> <table class='table table-bordered '>";
                                        string m_SectionHtml = "";
                                        Int32 m_ItemIndex = 1;
                                        Int32 m_TDIndex = 1;
                                        Int32 m_AllItemCount = m_dtSouce.Columns.Count;

                                        Dictionary<string, string> m_QueryCtrls = new Dictionary<string, string>(); ;
                                        if (p_QueryID != "")
                                        {
                                            foreach (string m_QueryCtrl in p_QueryID.Split(';'))
                                            {
                                                if (m_QueryCtrl != "")
                                                {
                                                    if (m_QueryCtrl.Split(',').Length != 0)
                                                    {
                                                        string m_QueryID = m_QueryCtrl.Split(',')[0];//QueryID
                                                        string m_PutAttr = m_QueryCtrl.Split(',')[1];//放查詢的欄位
                                                        string m_QueryParam = "";//帶入參數
                                                        foreach (string QueryParam in m_QueryCtrl.Split(',')[2].ToString().Split('+'))//帶入參數
                                                        {
                                                            foreach (DataRow EvRow in m_dtSouce.Rows)
                                                            {
                                                                m_QueryParam += (m_QueryParam == "" ? "" : "|") + QueryParam + ":" + EvRow[QueryParam].ToString();
                                                            }
                                                        }
                                                        m_OnClick = " onclick='OpenQueryForm(\"" + m_QueryMode + "Qry" + m_PutAttr + "\",\"" + p_ControlType + "\",\"" + m_QueryID + "\",\"\",\"" + m_QueryParam + "\",\"" + m_QueryMode + "\")'";
                                                        m_QueryCtrls[m_PutAttr] = "<img name='" + m_QueryMode + "Qry" + m_PutAttr + "' class='hand' " + m_OnClick + " src='Img/Query.png' alt='Query' height='20' width='20'>";
                                                    }
                                                }
                                            }
                                        }

                                        foreach (DataRow EvRow in m_dtSouce.Rows)
                                        {
                                            string m_FieldID = "";
                                            foreach (DataColumn Column in m_dtSouce.Columns)
                                            {
                                                m_FieldID = Column.ColumnName;
                                                string m_ControlValue = EvRow[m_FieldID].ToString();
                                                if (m_QueryCtrls.ContainsKey(m_FieldID))
                                                {
                                                    m_ControlValue = m_QueryCtrls[m_FieldID].ToString() + EvRow[m_FieldID].ToString();
                                                }
                                                bool m_isLastItem = m_AllItemCount == m_ItemIndex;
                                                string m_LanguageValue = GetMultiLanguage(m_FieldID, p_Language, p_Employee);
                                                if (Array.Exists(m_hidProps, element => element == m_FieldID))
                                                {
                                                    m_hidCtrls += string.Format(m_hidProp, m_FieldID, EvRow[m_FieldID].ToString());
                                                }
                                                string m_class = "bg-info text-right";
                                                m_SectionHtml = GeneratorTR(m_SectionHtml, p_SectionType, "", m_ControlValue, m_LanguageValue, m_FieldID, m_ColumnCount, "", "", m_isLastItem, m_ItemIndex, m_TDIndex, m_class);
                                                m_ItemIndex++;
                                                m_TDIndex++;
                                            }
                                        }
                                        m_TableHtml += m_SectionHtml;
                                        m_TableHtml += " </table></div>";
                                    }
                                    else if (p_ControlType == "TitleGrid")
                                    {
                                        string m_Height = "style='height:200px;'";
                                        if (p_Height != "")
                                        {
                                            m_Height = "style='height:" + p_Height + ";'";
                                        }
                                        string m_TableKey = p_ControlRef.Split(';')[0];
                                        string m_TableColumn = p_ControlRef.Split(';')[1];
                                        string m_TableWidth = p_ControlRef.Split(';')[2];
                                        object[] m_Params = new object[2];
                                        m_Params[0] = p_Employee;
                                        Dictionary<string, object> m_dicParam = new Dictionary<string, object>();
                                        if (m_TableKey != "")
                                        {
                                            foreach (string Key in m_TableKey.Split(','))
                                            {
                                                if(p_DataTable.Select().Count() > 0 ) //casper add for 表單新增時需秀grid
                                                    m_dicParam.Add(Key, p_DataTable.Rows[0][Key]);
                                            }
                                        }
                                        m_Params[1] = m_dicParam;
                                        DataTable m_dtSouce = GetQueryMethod(p_QueryMethod, m_Params, p_Employee);

                                        string m_TitleGrid = "<div class='TitleGridDivHead'><table class='table table-bordered table-hover MaintainTbHead'>";
                                        string m_GridTH = "{0}<tr class='bg-light'>";
                                        Int32 m_ColCount = 0;
                                        foreach (string Column in m_TableColumn.Split(','))
                                        {
                                            m_GridTH += "<th style='width:" + m_TableWidth.Split(',')[m_ColCount] + "%'>";
                                            m_GridTH += "<label id='lbl" + Column + "'>" + GetMultiLanguage(Column, p_Language, p_Employee) + "</label>";
                                            m_GridTH += "</th>";
                                            m_ColCount++;
                                        }
                                        m_GridTH += "</tr></table></div> ";
                                        string m_Reflash = " <img class='imgReflash' QueryMethod='" + p_QueryMethod + "' TableKey='" + m_TableKey + "' TableColumn='" + m_TableColumn + "' TableWidth='" + m_TableWidth + "'  id='ref" + p_ControlType + m_ControlID + "' alt='refresh' onclick='refTitleGrid(this,\"" + p_ControlType + m_ControlID + "\");' src='Img/refresh.gif'>";
                                        string m_GridData = "<div class='scroll gd" + p_ControlType + m_ControlID + " TitleGridDivBody ' " + m_Height + "><table  data='{0}'  id='tb" + p_ControlType + m_ControlID + "' class='table table-bordered table-hover MaintainTbBody'>";
                                        string m_PageCtrl = "";
                                        string m_JSON = "";
                                        m_JSON = JsonSerializer(m_dtSouce).Replace(@"'", "");

                                        if (m_dtSouce.Rows.Count > 50)
                                        {
                                            m_PageCtrl = "<tr><th>" + m_Reflash + "</th><th><span  class='hand' title='Prev' onclick='TitleGridPrev(\"" + p_ControlType + m_ControlID + "\",\"" + m_TableColumn + "\",\"" + m_TableWidth + "\")'>　＜＜　</span> <span title='Next'  class='hand' onclick='TitleGridNext(\"" + p_ControlType + m_ControlID + "\",\"" + m_TableColumn + "\",\"" + m_TableWidth + "\")'>　＞＞　</span></th><th id='thPageCtrl" + m_ControlID + "' colspan='" + (m_TableColumn.Split(',').Length - 2).ToString() + "'><span id='PageInfo" + p_ControlType + m_ControlID + "' PageIndex='1'> 1 ～ 50 </span>/<span  id='GridCount" + p_ControlType + m_ControlID + "'>" + m_dtSouce.Rows.Count.ToString() + "</span></th></tr>";
                                        }
                                        else
                                        {
                                            m_PageCtrl = "<tr><th>" + m_Reflash + "</th><th id='thPageCtrl" + m_ControlID + "' colspan='" + (m_TableColumn.Split(',').Length - 1).ToString() + "'><span id='PageInfo" + p_ControlType + m_ControlID + "' PageIndex='1'> 1 ～ " + m_dtSouce.Rows.Count.ToString() + " </span>/<span  id='GridCount" + p_ControlType + m_ControlID + "'>" + m_dtSouce.Rows.Count.ToString() + "</span></tr>";
                                        
                                        }
                                        m_GridTH = string.Format(m_GridTH, m_PageCtrl);
                                        
                                        m_GridData = string.Format(m_GridData, m_JSON);
                                        DataRow[] m_GridRows = m_dtSouce.Rows.Cast<System.Data.DataRow>().Take(50).ToArray();
                                           
                                       
                                        foreach (DataRow dr in m_GridRows)//dt.Rows.Cast<System.Data.DataRow>().Take(n)
                                        {
                                            m_GridData += "<tr>";
                                            m_ColCount = 0;
                                            foreach (string Column in m_TableColumn.Split(','))
                                            {
                                                m_GridData += "<td style='width:" + m_TableWidth.Split(',')[m_ColCount] + "%'>";
                                                m_GridData += dr[Column].ToString();
                                                m_GridData += "</td>";
                                                m_ColCount++;
                                            }
                                            m_GridData += "</tr>";
                                        }
                                        m_TitleGrid += m_GridTH + m_GridData;
                                        m_TitleGrid += " </table></div>";
                                        m_TableHtml += m_TitleGrid;
                                    }
                                    m_TableHtml += m_hidCtrls + " </div> </div>";
                                }
                                m_Control = m_TableHtml;
                            }
                            break;
                        case "Date":
                        case "DateTime":
                        case "Text":
                            string m_SelTime = "";
                            string m_SelTimeVal = "";
                            if (p_ControlType == "DateTime")
                            {
                                m_Options = "<option value='08' >08:00</option>";
                                m_Options += "<option value='09' >09:00</option>";
                                m_Options += "<option value='10' >10:00</option>";
                                m_Options += "<option value='11' >11:00</option>";
                                m_Options += "<option value='12' >12:00</option>";
                                m_Options += "<option value='13' >13:00</option>";
                                m_Options += "<option value='14' >14:00</option>";
                                m_Options += "<option value='15' >15:00</option>";
                                m_Options += "<option value='16' >16:00</option>";
                                m_Options += "<option value='17' >17:00</option>";
                                m_Options += "<option value='18' >18:00</option>";
                                m_Options += "<option value='19' >19:00</option>";

                                if (p_DataTable != null)
                                {
                                    if (p_DataTable.Rows.Count != 0)
                                    {
                                        m_SelTimeVal = p_DataTable.Rows[0][p_FieldID.Replace("Date", "Time")].ToString();
                                    }
                                    m_Options = m_Options.Replace("value='" + m_SelTimeVal + "'", "value='" + m_SelTimeVal + "' selected");
                                }
                                m_SelTime = " <select style='width: 100px;display:inline;' " + p_ReadOnly + " mode='" + m_QueryMode + "' id='" + m_QueryMode + "sel" + m_ControlID.Replace("Date","Time") + "'  > " + m_Options + "</select>";
                            }

                            if (p_ControlRef != "")
                            {
                                string[] m_txtDesc = { p_ControlRef };
                                SysEntity.TransResult m_TransResultTitle = g_BL_System.GetListMultiLanguage(p_Employee, p_Language, m_txtDesc);

                                Dictionary<string, Object> m_distxtDesc = new Dictionary<string, object>();
                                if (m_TransResultTitle.isSuccess)
                                {
                                    m_distxtDesc = (Dictionary<string, Object>)m_TransResultTitle.ResultEntity;
                                }

                                m_Control = "<label Remind='Y' class='col-lg-2 control-label' id='lbl" + p_ControlRef + "'  style='color:red' >" + m_distxtDesc[p_ControlRef].ToString() + "</label></br>";
                            }
                            if (p_ControlType != "Text")
                            {
                                m_style = "style='width:120px;display:inline; '";
                            }

                            m_Control += "<input class='form-control'  " + m_ModifyKey + "  id='" + m_QueryMode + "txt" + m_ControlID + "' " + m_DataValue + "  " + m_TextType + " " + p_ReadOnly + " " + p_Valid + "  " + m_style + " " + m_MaxLength + "  type='text' />" + m_SelTime;
                            break;
                        case "MultiText":
                            if (p_Height != "")
                            {
                                m_style = "style='width:90%;Height:150px'";
                            }
                            else
                            {
                                m_style = "style='width:90%;Height:" + p_Height + "'";
                            }
                            m_Control = "  <textarea class='form-control'  " + m_ModifyKey + "  id='" + m_QueryMode + "txt" + m_ControlID + "' " + p_ReadOnly + "  " + p_Valid + " " + m_style + "  " + m_MaxLength + "  >" + m_DataValue + "</textarea> ";
                            break;
                        case "Query":
                            m_OnChange = " onchange='QueryOnChange(\"" + m_QueryMode + "Qry" + m_ControlID + "\",\"" + m_QueryMode + "\")' ";
                            string m_lblDesc = "";
                            if (p_DataTable != null)
                            {
                                if (p_DataTable.Rows.Count != 0)
                                {
                                    //m_lblDesc = "<label id='lbl{0}'>{1}</label>";
                                    string m_EventID = g_BL_System.GetEventIDByQueryID(p_Employee, p_QueryID);
                                    if (m_EventID != "")
                                    {
                                        //1 SysEntity.Employee  2 Dictionary<string, Object> p_dic, 3 string p_PageIndex, 4 string p_PageSize, 5 Dictionary<string, Object> p_Param, 6 Boolean p_isDataOnly 
                                        Dictionary<string, Object> m_dic = new Dictionary<string, object>();
                                        m_dic[p_FieldID] = m_OnlyValue;
                                        object[] m_Params = new object[6];
                                        m_Params[0] = p_Employee;
                                        m_Params[1] = m_dic;
                                        m_Params[2] = null;
                                        m_Params[3] = null;
                                        m_Params[4] = null;
                                        m_Params[5] = true;

                                        DataTable m_dtSouce = GetQueryMethod(m_EventID, m_Params, p_Employee);

                                        if (p_ControlRef != "")
                                        {
                                            string[] m_ColumnRef = p_ControlRef.Split(',');
                                            foreach (string Column in m_ColumnRef)
                                            {
                                                if (p_FieldID != Column)
                                                {
                                                    if (m_dtSouce.Rows.Count != 0)
                                                    {
                                                        if (!p_DataTable.Columns.Contains(Column))
                                                        {
                                                            m_lblDesc += string.Format("<label id='lbl{0}'>{1}</label>", Column, m_dtSouce.Rows[0][Column].ToString());
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                            if (m_DisDefVal.Count != 0)
                            {
                                if (p_ControlRef != "")
                                {
                                    string[] m_ColumnRef = p_ControlRef.Split(',');
                                    foreach (string Column in m_ColumnRef)
                                    {
                                        if (m_DisDefVal.ContainsKey(Column))
                                        {
                                            if (Column == p_FieldID)
                                            {
                                                m_DataValue =" value='"+ m_DisDefVal[Column].ToString()+"'";
                                            }
                                            else
                                            {
                                                m_lblDesc += string.Format("<label id='lbl{0}'>{1}</label>", Column, m_DisDefVal[Column].ToString());
                                            
                                            }
                                        }
                                    
                                    }
                                }

                            }

                            m_OnClick = " onclick='OpenQueryForm(\"" + m_QueryMode + "Qry" + m_ControlID + "\",\"" + p_ControlType + "\",\"" + p_QueryID + "\",\"" + p_ControlRef + "\",\"\",\"" + m_QueryMode + "\")'";
                            string m_img = "<img name='" + m_QueryMode + "Qry" + m_ControlID + "' class='hand' " + m_OnClick + " src='Img/Query.png' alt='Query' height='20' width='20'/>";
                            if (p_ReadOnly.IndexOf("ReadOnly='true'") != -1)
                            {
                                m_img = "";
                            }
                            m_Control = "<table><tr><td  class='qry'><input class='form-control'  " + m_ModifyKey + "  " + m_OnChange + "  FieldID='" + p_FieldID + "' id='" + m_QueryMode + "Qry" + m_ControlID + "' " + m_DataValue + "  " + m_TextType + " " + p_ReadOnly + " " + p_Valid + " style='width:100%'  " + m_MaxLength + "  type='text' /></td><td>" + m_img + m_lblDesc + "</td></tr></table>";
                            break;
                        case "Button":
                            string m_Class = "class='btn btn-primary mb10 mr5  notification'";
                            switch (p_FieldID)
                            {
                                case "Delete":
                                case "Reject":
                                    m_Class = "class='btn btn-system mb10 mr5 notification'";
                                    break;
                            }
                            if (p_FieldID == "ExportExcel")
                            {
                                m_Control = "<input  style=\"background-image:url('Img/ExpXls.png') ;background-repeat:no-repeat;background-position:center;width:50px\" title=\"Export Excel\" " + m_Class + "  id='xlsbtn" + m_ControlID + "'  " + p_ReadOnly + "   onclick='ExportExcel(\"" + p_EventID + "\",\"" + p_SiteFormID + "\",\"" + p_ControlRef + "\")'  type='button' />";
                            }
                            else if (p_FieldID == "BatchApprove")
                            {
                                m_Control = "<input " + m_Class + " id='btn" + p_FieldID + "' onclick='OnClickCommonButton(\"" + p_EventID + "\",\"" + p_FieldID + "\",\"BatchApprove\",\"" + p_LanguageValue + "\",\"" + p_ControlRef + "\",\"" + p_SiteFormID + "\")' type='button' value='" + p_LanguageValue + "' />";
                            }
                            else if (p_FieldID == "BatchReject")
                            {
                                m_Control = "<input " + m_Class + " id='btn" + p_FieldID + "' onclick='OnClickCommonButton(\"" + p_EventID + "\",\"" + p_FieldID + "\",\"BatchReject\",\"" + p_LanguageValue + "\",\"" + p_ControlRef + "\",\"" + p_SiteFormID + "\")' type='button' value='" + p_LanguageValue + "' />";
                            }
                            else if (p_FieldID == "Search" || p_FieldID == "BatchSearch")
                            {
                                m_Control = "<input " + m_Class + ((p_Width == "display:none") ? m_style : "") + "  id='btn" + m_ControlID + "'  " + p_ReadOnly + "  " + m_ControlRef + " ondblclick='return false'   onclick='ActionModel(\"" + p_EventID + "\",\"" + p_FieldID + "\",\"" + p_EventActionType + "\");InitSearchInfo(\"" + p_EventID + "\",\"" + p_FieldID + "\",\"" + p_EventActionType + "\",\"\")' type='button' value='" + p_LanguageValue + "' />";
                            }
                            else
                            {

                                if (m_ControlID == "History")
                                {
                                    if (p_DataTable != null)
                                    {
                                        if (p_DataTable.Rows.Count != 0)
                                        {
                                            var m_SiteFormID = p_DataTable.Rows[0]["SiteFormID"].ToString();
                                            string m_AllIetm = "";

                                            if (p_EventID != "")
                                            {
                                                Dictionary<string, object> m_dicParam = new Dictionary<string, object>();
                                                m_dicParam["SiteFormID"] = m_SiteFormID;
                                                object[] m_Params = new object[2];
                                                m_Params[0] = p_Employee;
                                                m_Params[1] = m_dicParam;
                                                DataTable m_dtSouce = GetQueryMethod(p_EventID, m_Params, p_Employee);
                                                string m_HistoryOnClick = " onclick='HistoryOnClick(this,\"{3}\")' ";
                                                string m_Ietm = " <li " + m_HistoryOnClick + "><table class='tbHisItem'><tr><td style='width:25%;'>{0}</td><td style='width:20%;'>{1}</td><td style='width:55%;'>{2}</td></tr></table>";
                                                m_AllIetm += string.Format(m_Ietm, "User", "Type", "Time", "") + "</li><li class='divider'></li>";
                                                foreach (DataRow dr in m_dtSouce.Rows)
                                                {
                                                    m_AllIetm += string.Format(m_Ietm, dr["ActionUser"].ToString(), dr["ActionType"].ToString(), dr["ActionTime"].ToString(), dr["TempKey"].ToString());
                                                }
                                            }
                                            m_Control = "<div class='btn-group'>  <input id='hidOldTempKey'  type='hidden' /> <button  type='button' class='btn btn-info dropdown-toggle'  data-toggle='dropdown' id='btn" + m_ControlID + "'  " + p_ReadOnly + " >" + p_LanguageValue + "&nbsp;&nbsp;<span class='caret'></span></button><input style='display:none' class='btn btn-danger ml5 notification' id='btnRollback'  ondblclick='return false'  onclick='ActionModel(\"" + p_QueryMethod + "\",\"" + p_FieldID + "\",\"Rollback\")' type='button' value='Rollback' /> ";
                                            m_Control += "<ul class='dropdown-menu divHistoryUl' role='menu' >";
                                            m_Control += m_AllIetm;
                                            m_Control += "</ul></div>";
                                            m_Control += "<table class='tbHisItem' style='display:none' id='tbSelectedLi'><tr><td id='selLi'></td></tr><tr><td>";
                                            m_Control += "</td></tr></table>";
                                        }
                                    }

                                }
                                else
                                {
                                    string m_GridOndblclick = "";
                                    if (p_FieldID == "Query")
                                    {
                                        m_GridOndblclick = " EventID='" + p_EventID + "'" + " FieldID='" + p_FieldID + "'" + " EventActionType='" + p_EventActionType + "'" + " LanguageValue='" + p_LanguageValue + "'" + " QueryMethod='" + p_QueryMethod + "'" + " EventFollow='" + p_EventFollow + "'";
                                    }
                                    if (p_ControlRef == "onScript")
                                    {
                                        m_Control = "<input " + m_Class + "  id='" + m_QueryMode + "btn" + m_ControlID + "' onclick='onScript(\"" + p_FieldID + "\")' type='button' value='" + p_LanguageValue + "' />";

                                    }
                                    else
                                    {
                                        m_Control = "<input " + m_GridOndblclick + " " + m_Class + "  id='" + m_QueryMode + "btn" + m_ControlID + "' " + m_ControlRef + "  onclick='OpenMaintain(\"" + p_EventID + "\",\"" + p_FieldID + "\",\"" + p_EventActionType + "\",\"" + p_LanguageValue + "\",\"" + p_QueryMethod + "\",\"" + p_EventFollow + "\")' type='button' value='" + p_LanguageValue + "' />";
                                    }
                                }
                            }
                            break;
                        case "Display":
                        case "Label"://label
                            if (p_ControlType == "Display")
                            {
                                if (p_EventID != "")
                                {
                                    m_OnClick = "";
                                }
                                m_Control = "<label class='col-lg-2 control-label'  FieldID='" + p_FieldID + "'" + "  id='" + m_QueryMode + "dis" + m_ControlID + "'  " + m_style + " >" + m_DataValue + "</label>";
                            }
                            else
                            {
                                m_Control = "<label class='col-lg-2 control-label'  id='lbl" + m_ControlID + "'  " + m_style + " >" + m_DataValue + "</label>";
                            }
                            break;
                        case "Select":
                            string m_DefVal = "";
                            if (p_EventID != "")
                            {//GetEventSettings(SysEntity.Employee p_EmployeeEntity, Dictionary<string, Object> 
                                Dictionary<string, Object> m_EventID=new Dictionary<string,object>();
                                m_EventID["EventID"] = p_EventID;
                                SysEntity.TransResult m_EventTransResult = g_BL_System.GetEventSettings(p_Employee, m_EventID);
                                if (m_EventTransResult.isSuccess)
                                {
                                    if (((DataTable)m_EventTransResult.ResultEntity).Rows.Count != 0)
                                    {
                                        m_OnChange = " onchange='SelectOnChange(this,\"" + p_EventID + "\",\"" + p_ControlRef.Split(';')[1] + "\",\"" + p_EventID + "\",\"" + m_QueryMode + "\")' ";
                                    }
                                    else
                                    {
                                        m_OnChange = "onchange='" + p_EventID + "'";
                                    }
                                }
                            }

                            if (p_ControlRef != "" )
                            {
                                m_selText = p_ControlRef.Split(';')[0].ToString().Split(',')[1];
                                m_selValue = p_ControlRef.Split(';')[0].ToString().Split(',')[0];

                            }
                            if (p_QueryMethod != "")
                            {
                                object[] m_Params = new object[2];
                                m_Params[0] = p_Employee;
                                m_Params[1] = null;
                                DataTable m_dtSouce = GetQueryMethod(p_QueryMethod, m_Params, p_Employee);
                                if (p_DefaultValue == "null")
                                {
                                    m_DefVal = "DefVal='null'";
                                    m_Options += "<option value='' " + m_Selected + "></option>";

                                }
                                if (p_SectionID == "divQueryArea" && p_DefaultValue == "ALL")
                                {
                                    m_Options += "<option value='' " + m_Selected + ">ALL</option>";
                                }
                               
                                foreach (DataRow EvRow in m_dtSouce.Rows)
                                {
                                    m_Selected = "";
                                    if (EvRow[m_selValue].ToString() == m_DataValue)
                                    {
                                        m_Selected = "selected";
                                    }
                                    m_Options += "<option value='" + EvRow[m_selValue].ToString() + "' " + m_Selected + ">" + EvRow[m_selText].ToString() + "</option>";
                                }
                            }
                            else
                            {
                                if (p_ControlRef.IndexOf(",") != -1)
                                {
                                    foreach (string m_option in p_ControlRef.Split(','))
                                    {
                                        m_Options += "<option value='" + m_option + "' " + m_Selected + ">" + m_option + "</option>";
                                    }
                                }
                            }
                            m_Control = " <select selQID='" + p_QueryMethod + "' " + m_DefVal + " " + m_ModifyKey + "  selText='" + m_selText + "' " + m_style.Replace("90%;", "auto;") + " " + p_ReadOnly + " mode='" + m_QueryMode + "' id='" + m_QueryMode + "sel" + m_ControlID + "' " + p_Valid + " " + m_OnChange + "  > " + m_Options + "</select>";
                            break;

                        case "Radio":
                            string m_Radios = "";
                            string m_RadVal = "";
                            if (p_QueryMethod != "")
                            {
                                object[] m_Params = new object[2];
                                m_Params[0] = p_Employee;
                                m_Params[1] = null;
                                DataTable m_dtSouce = GetQueryMethod(p_QueryMethod, m_Params, p_Employee);
                                string m_Text = p_ControlRef.Split(',')[1];
                                string m_Value = p_ControlRef.Split(',')[0];
                                foreach (DataRow EvRow in m_dtSouce.Rows)
                                {
                                    m_RadVal = "";
                                    if (EvRow[m_Value].ToString() == m_DataValue)
                                    {
                                        m_RadVal = "checked";
                                    }
                                    m_Radios += "<td><input type='radio' name='rod" + m_ControlID + "' " + m_RadVal + "  value='" + EvRow[m_Value].ToString() + "'/></td><td><label id='lbl" + m_Value + "' >" + EvRow[m_Text].ToString() + "</label></td>";
                                }
                            }
                            m_Control = "<table id='" + m_QueryMode + "tabrdo" + m_ControlID + "'><tr>" + m_Radios + "</tr></table>";

                            break;
                        case "CheckBox":
                            string m_CheckedVal = "";
                            string m_Check = "";
                            if (p_QueryMethod != "")
                            {
                                object[] m_Params = new object[2];
                                m_Params[0] = p_Employee;
                                m_Params[1] = null;
                                DataTable m_dtSouce = GetQueryMethod(p_QueryMethod, m_Params, p_Employee);
                                string m_Text = p_ControlRef.Split(',')[1];
                                string m_Value = p_ControlRef.Split(',')[0];
                                foreach (DataRow EvRow in m_dtSouce.Rows)
                                {
                                    m_CheckedVal = "";
                                    if (m_DataValue != "")
                                    {
                                        foreach (string m_val in m_DataValue.Split(','))
                                        {
                                            if (m_val == EvRow[m_Value].ToString())
                                            {
                                                m_CheckedVal = "checked";
                                                break;
                                            }
                                        }
                                    }
                                    m_Check += "<td><input type='checkbox' name='chk" + m_ControlID + "' " + m_CheckedVal + "  value='" + EvRow[m_Value].ToString() + "'/></td><td><label id='lbl" + m_Value + "' >" + EvRow[m_Text].ToString() + "</label></td>";
                                }
                                m_Check = "<table id='" + m_QueryMode + "tabchk" + m_ControlID + "'><tr>" + m_Check + "</tr></table>";
                            }
                            else
                            {
                                if (m_DataValue == "1")
                                {
                                    m_CheckedVal = "checked";
                                }
                                m_Check = "<input id='" + m_QueryMode + "chk" + m_ControlID + "'  " + m_CheckedVal + "  value='1' type='checkbox' />";
                            }
                            m_Control += m_Check;
                            break;
                    }
                }
                else if (p_SectionType == "Detail")
                {
                    string m_DivHeigh = "height:450px";
                    if (p_Height != "")
                    {
                        m_DivHeigh = "height:" + p_Height;
                    }

                    string m_DetailDiv = "";

                    string[] m_Title = { p_FieldID };
                    SysEntity.TransResult m_TransResultTitle = g_BL_System.GetListMultiLanguage(p_Employee, p_Language, m_Title);

                    Dictionary<string, Object> m_disTitle = new Dictionary<string, object>();
                    if (m_TransResultTitle.isSuccess)
                    {
                        m_disTitle = (Dictionary<string, Object>)m_TransResultTitle.ResultEntity;
                    }


                    string m_DetailTitle = "<label id='lbl" + p_FieldID + "'>" + m_disTitle[p_FieldID].ToString() + "</label>";

                    m_DetailDiv += "<div class='Dtl-heading' id='DtlPanel' style='width:1024px;background:#bce4f2;max-width:1400px;min-width:800px;'>";
                    m_DetailDiv += "<div style='width:50%;text-align:left;float: left;'>&nbsp;&nbsp;&nbsp;<img src='Img/DtlIco.png' alt='DtlIco' height='16' width='19'>" + m_DetailTitle + "</div><div style='width:50%;text-align:right;float:right ;'><i class='fa fa-plus-circle hand' aria-hidden='true' title='Add Rows'  onclick='AddRows(\"DtlTable" + p_FieldID + "\",\"" + p_ActionMode + "\")' alt='AddRows' ></i>&nbsp;&nbsp;&nbsp;</div></div>";
                    m_DetailDiv += "<div class='table-responsive' style='width:100%'>";


                    DataTable m_dtDetailItems = g_BL_System.GetViewDetailItems(p_Employee, p_SiteFormID, "divMaintain", p_FieldID);


                    string[] m_Keys = new string[m_dtDetailItems.Rows.Count];
                    Int32 m_ColumnCount = 0;
                    foreach (DataRow m_dr in m_dtDetailItems.Rows)
                    {
                        m_Keys[m_ColumnCount] = m_dr["FieldID"].ToString();
                        m_ColumnCount++;
                    }
                    Int32 m_daicWidth = 1100;
                    Int32 m_Incre = 165;
                    if (m_dtDetailItems.Rows.Count > 10)
                    {
                        m_daicWidth = m_daicWidth + (m_Incre * (m_dtDetailItems.Rows.Count - 10));
                    }

                    Dictionary<string, Object> m_Languages = new Dictionary<string, object>();

                    SysEntity.TransResult m_TransResult = g_BL_System.GetListMultiLanguage(p_Employee, p_Language, m_Keys);
                    if (m_TransResult.isSuccess)
                    {
                        m_Languages = (Dictionary<string, Object>)m_TransResult.ResultEntity;
                    }

                    string m_HeadTable = "<div class='scroll' id='DtlHeadScroll' style='height:55px;width:1024px;max-width:1400px;min-width:800px;'><table id='DtlHeadTable' class='tabGridHead table table-striped'  style='height:50px;min-width:" + m_daicWidth.ToString() + "px ;width:100%' id='tab" + m_ControlID + "'><tr>{0}</tr></table></div>";
                    string m_HeadTableTH = "";
                    string m_BodyTable = "<div class='scroll' id='DtlScroll' style='" + m_DivHeigh + ";width:1024px;max-width:1400px;min-width:800px;'><table Layout='{0}' id='DtlTable" + p_FieldID + "' class='DtlTable table table-bordered' style='min-width:" + m_daicWidth.ToString() + "px;height:auto;width:100%'>{1}</table></div>";
                    string m_BodyTableTR = "";
                    string m_BodyTableAllTR = "";
                    string m_BodyDefaultRow = "";
                    Int32 m_DataAllCount = 0;
                    Int32 m_DataCount = 1;
                    DataTable m_ValueTable = new DataTable();

                    m_BodyDefaultRow = "id='DtlDefRow' style='display:none'";//預設一筆隱藏TR 用於AddRow
                    m_BodyTableTR += GetDtlTable(p_Employee, m_dtDetailItems, null, m_BodyDefaultRow, p_SectionID, "0", p_ActionMode, p_SectionID, p_Language, m_Languages, ref m_HeadTableTH, p_FieldID, true);


                    if (p_ActionMode == "Add")
                    {


                    }
                    else
                    {
                   
                        //Detail 明細資料(主從檔根據QueryMethod及Key查詢後帶出)
                        if (p_QueryMethod != "")
                        {
                            if (p_ControlRef != "")
                            {
                                string m_Where = "";
                                Int32 m_ParamCount = 1;
                                object[] m_DtlParams = null;
                                string[] m_arrControlRef = null;
                                if (p_ControlRef.IndexOf(';') != -1)
                                {
                                    string m_DtlWhereParams = p_ControlRef.Split(';')[1].ToString();
                                    m_DtlParams = new object[p_ControlRef.Split(';')[0].Split(',').Length + 1];
                                    m_arrControlRef = p_ControlRef.Split(';')[0].Split(',');
                                    m_Where = m_DtlWhereParams + "='" + p_FieldID + "'";
                                }
                                else
                                {
                                    m_DtlParams = new object[p_ControlRef.Split(',').Length + 1];
                                    m_arrControlRef = p_ControlRef.Split(',');
                                    m_DtlParams[0] = p_Employee;
                                }
                                m_DtlParams[0] = p_Employee;

                                foreach (string Param in m_arrControlRef)
                                {
                                    if (p_DataTable != null)
                                    {
                                        if (p_DataTable.Rows.Count != 0)
                                        {
                                            foreach (DataRow dr in p_DataTable.Select(m_Where))
                                            {
                                                m_DtlParams[m_ParamCount] = dr[Param].ToString();
                                            }
                                        }
                                    }
                                  
                                    m_ParamCount++;
                                }

                                DataTable m_dtSouce = GetQueryMethod(p_QueryMethod, m_DtlParams, p_Employee);

                                m_DataAllCount = m_dtSouce.Rows.Count;
                                foreach (DataRow dr in m_dtSouce.Rows)
                                {
                                    m_BodyDefaultRow = "";
                                    if (m_DataCount == 1)
                                    {
                                        m_ValueTable = m_dtSouce.Copy();
                                    }
                                    m_ValueTable.Clear();
                                    DataRow m_Valdr = m_ValueTable.NewRow();
                                    foreach (DataColumn m_dc in m_ValueTable.Columns)
                                    {
                                        m_Valdr[m_dc.ColumnName] = dr[m_dc.ColumnName];
                                    }
                                    m_ValueTable.Rows.Add(m_Valdr);
                                    m_BodyTableTR += GetDtlTable(p_Employee, m_dtDetailItems, m_ValueTable, m_BodyDefaultRow, p_SectionID, m_DataCount.ToString(), p_ActionMode, p_SectionID, p_Language, m_Languages, ref m_HeadTableTH, p_FieldID);
                                    m_DataCount++;
                                }


                            }
                        }
                        else
                        {   //Detail 明細資料(單檔多筆)
                            m_DataAllCount = p_DataTable.Rows.Count;
                            foreach (DataRow dr in p_DataTable.Rows)
                            {
                                m_BodyDefaultRow = "";
                                if (m_DataCount == 1)
                                {
                                    m_ValueTable = p_DataTable.Copy();
                                }
                                m_ValueTable.Clear();
                                DataRow m_Valdr = m_ValueTable.NewRow();
                                foreach (DataColumn m_dc in m_ValueTable.Columns)
                                {
                                    m_Valdr[m_dc.ColumnName] = dr[m_dc.ColumnName];
                                }
                                m_ValueTable.Rows.Add(m_Valdr);
                                m_BodyTableTR += GetDtlTable(p_Employee, m_dtDetailItems, m_ValueTable, m_BodyDefaultRow, p_SectionID, m_DataCount.ToString(), p_ActionMode, p_SectionID, p_Language, m_Languages, ref m_HeadTableTH, p_FieldID);
                                m_DataCount++;
                            }
                        }

                    }





                    m_BodyTableAllTR += m_BodyTableTR;
                    string m_Layout = JsonSerializer(m_dtDetailItems);
                    m_BodyTable = string.Format(m_BodyTable, m_Layout, m_BodyTableAllTR);
                    m_HeadTable = string.Format(m_HeadTable, "<th class='DelChk'><i class='fa fa-trash' aria-hidden='true'></i></th><th id='DtlTHSeq' >SEQ</th>" + m_HeadTableTH);
                    m_DetailDiv += m_HeadTable + m_BodyTable + " </div>";

                    m_Control = "<div id='DtlTD" + p_FieldID + "' field='" + p_FieldID + "'  class='DtlTD' style='scroll:auto;max-width:1400px;min-width:800px;'>" + m_DetailDiv + " </div>";
                }
               
            }
            catch (Exception ex)
            {
                throw ex;
            }
            return m_Control;
        }
        
        public string GetDtlTable(SysEntity.Employee p_Employee, DataTable p_dtDetail, DataTable p_ValueTable, string p_BodyDefaultRow, string p_SectionID, string p_DtlRowCount, string p_ActionMode, string p_SiteFormID, string p_Language, Dictionary<string, Object> p_Languages, ref string p_HeadTableTH,string p_FieldID,bool p_isNewRow=false)
        {
            string m_BodyTableAllTR = "";
            string m_BodyDefaultRow = p_BodyDefaultRow;
            string m_BodyTableTR = "<tr class='DtlRows' " + m_BodyDefaultRow + " RowAction='Edit'>{0}</tr>";
            string m_BodyTableTD = "";
            string m_HeadTableTH = "";
            string m_Key = "";
            try
            {
                foreach (DataRow m_dr in p_dtDetail.Rows)
                {
                    string m_DtlControlID = m_dr["FieldID"].ToString() + p_DtlRowCount.ToString();
                    if (m_BodyDefaultRow != "")
                    {
                        m_DtlControlID = m_dr["FieldID"].ToString() + "0";
                    }
                    else
                    {
                        m_DtlControlID = m_dr["FieldID"].ToString() + p_DtlRowCount.ToString();
                    }
                    string m_DtlControlType = m_dr["ControlType"].ToString();
                    string m_DtlSectionType = m_dr["SectionType"].ToString();
                    string m_DtlFieldID = m_dr["FieldID"].ToString();
                    string m_DtlQueryID = m_dr["QueryID"].ToString();
                    string m_DtlControlRef = m_dr["ControlRef"].ToString();
                    string m_DtlValidID = m_dr["ValidID"].ToString();
                    string m_DtlHeight = m_dr["Height"].ToString();
                    string m_DtlWidth = m_dr["Width"].ToString();
                    string m_DtlEventID = m_dr["EventID"].ToString();
                    string m_DtlEventAction = m_dr["EventAction"].ToString();
                    string m_DtlEventActionType = m_dr["EventActionType"].ToString();
                    string m_DtlValidSource = m_dr["ValidSource"].ToString();
                    string m_DtlValidParam = m_dr["ValidParam"].ToString();
                    string m_DtlValid = "";
                    string m_DtlMaxLength = m_dr["MaxLength"].ToString();
                    string m_DtlModifyKey = m_dr["ModifyKey"].ToString();
                    string m_DtlEventFollow = m_dr["EventFollow"].ToString();
                    string m_DtlQueryMethod = m_dr["QueryMethod"].ToString();
                    string m_DefaultValue = m_dr["DefaultValue"].ToString();

                    
                    string m_DtlReadOnly = "";
                    if (m_DtlWidth == "0")
                    {
                        m_DtlWidth = "display:none";
                    }
                    else
                    {
                        m_DtlWidth = "Width:" + m_DtlWidth + "%";
                    }
                    if (p_HeadTableTH == "")
                    {
                        m_HeadTableTH += "<th id='DtlTh" + m_DtlFieldID + "' style='" + m_DtlWidth + "' >" + p_Languages[m_DtlFieldID].ToString() + "</th>";
                    }

                    if (m_DtlModifyKey == "1")
                    {
                        m_Key = " RowKey='1' ";
                        if (!(p_isNewRow))
                        {
                            m_DtlReadOnly = " ReadOnly='true'  disabled='disabled' ";
                        }
                    }
                    else
                    {
                        m_Key = "";
                    }
                    if ((m_DtlModifyKey == "1" && p_ActionMode == "Add"))
                    {
                        m_Key = " RowKey='1' ";
                    }
                    if (m_DtlValidSource != "")
                    {
                        m_DtlValid = " ValidType='" + m_DtlValidSource + "'  onfocusout='" + m_DtlValidSource + "(this)' ";
                    }
                    m_BodyTableTD += "<td class='Dtltd' valuetype={0} " + m_Key + " Field='" + m_DtlFieldID + "' Value='{1}' style='" + m_DtlWidth + "' >{2}</td>";

             
                    string m_DtlControl = "";
                    string m_DtlValue = "";
                    string m_DtlValueType = "text";

                    if (p_ValueTable != null)
                    {
                        if (p_ValueTable.Rows.Count != 0)
                        {
                            m_DtlValue = p_ValueTable.Rows[0][m_DtlFieldID].ToString();
                            m_DtlControl = m_DtlValue;

                        }
                       
                    }
                    switch (m_DtlControlType)
                    {
                        case "Select":
                            break;
                        case "CheckBox":
                            string m_Check = "";
                            if (m_DtlValue == "1")
                            {
                                m_Check = "checked='checked'";
                            }
                            m_DtlControl = "<input disabled='disabled' name='" + m_DtlFieldID + "' " + m_Check + " type='checkbox' />";

                            break;
                        case "MultiSelect":
                            m_DtlValueType = "Multi";
                            break;
                    }


                    m_BodyTableTD = string.Format(m_BodyTableTD, m_DtlValueType, m_DtlValue, m_DtlControl);
                }
                if (p_HeadTableTH == "")
                {
                    p_HeadTableTH = m_HeadTableTH;
                }
                string m_DtlTRSeqIndex = p_DtlRowCount.ToString().PadLeft(3, '0');
                if (m_BodyDefaultRow != "")
                {
                    m_DtlTRSeqIndex = "000";
                }
                m_BodyTableTR = string.Format(m_BodyTableTR, "<td class='DelChk' ><input id='chkDtl'  style='width:20px'  value='1' type='checkbox' /><span class='glyphicon glyphicon-pencil' onclick='ModifyDetils(this,\"DtlTable" + p_FieldID + "\",\"" + p_ActionMode + "\")'></span></td><td Field='DtlTRSeq' class='RowIndex' Index='" + m_DtlTRSeqIndex + "' id='DtlTRSeq" + m_DtlTRSeqIndex + "'>" + m_DtlTRSeqIndex + "</td>" + m_BodyTableTD);
                m_BodyTableAllTR += m_BodyTableTR;
            }
            catch (Exception ex)
            {
                throw ex;
            }
            return m_BodyTableAllTR;
        }

        public DataTable GetParameter(string p_SectionType, string p_ParamID, SysEntity.Employee p_Employee)
        {
            DataTable m_dtReturn = new DataTable();
            try
            {
                SysEntity.Parameter m_Parameter = new SysEntity.Parameter();
                m_Parameter.ParamID = p_ParamID;
                m_Parameter.Text = p_SectionType;
                SysEntity.TransResult m_ParameterResult = g_BL_System.GetParameter(p_Employee, m_Parameter);
                if (m_ParameterResult.isSuccess)
                {
                    m_dtReturn = ((DataTable)m_ParameterResult.ResultEntity);
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            return m_dtReturn;
        }

        public Dictionary<string, Object> GetKeyByDic(string p_Keys, Dictionary<string, Object> p_DicParam)
        {
            Dictionary<string, Object> m_KeyDic = new Dictionary<string, object>();

            try
            {

                foreach (string p_Key in p_Keys.Split(','))
                {
                    m_KeyDic[p_Key] = p_DicParam[p_Key].ToString();
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            return m_KeyDic;
        }

        public DataTable GetQueryMethod(string p_QueryMethod, object[] p_Params, SysEntity.Employee p_Employee)
        {
            DataTable m_dtReturn = new DataTable();
            SysEntity.EventSetting m_EventEntry = new SysEntity.EventSetting();
            m_EventEntry.EventID = p_QueryMethod;

            try
            {
                SysEntity.TransResult m_TransResult = g_BL_System.GetEventSetting(p_Employee, m_EventEntry);
                string m_EventModel = "";
                string m_EventAction = "";
                string m_EventActionType = "";
                string m_EventRef = "";

                if (m_TransResult.isSuccess)
                {
                    foreach (DataRow dr in ((DataTable)m_TransResult.ResultEntity).Rows)
                    {
                        m_EventEntry.EventModel = dr["EventModel"].ToString();
                        m_EventEntry.EventAction = dr["EventAction"].ToString();

                        m_EventModel = dr["EventModel"].ToString();
                        m_EventAction = dr["EventAction"].ToString();
                        m_EventActionType = dr["EventActionType"].ToString();
                        m_EventRef = dr["EventRef"].ToString();
                    }
                }
                Type m_Type = GetBLType(m_EventModel);
                object instance = Activator.CreateInstance(m_Type);
                MethodInfo method = instance.GetType().GetMethod(m_EventAction);
                ParameterInfo[] m_Pars = method.GetParameters();
                object[] m_Params = new object[m_Pars.Length];
                
                if (m_Pars.Length > p_Params.Length)
                {
                    for (Int32 ParamCount = 0; ParamCount < m_Pars.Length; ParamCount++)
                    {
                        if (ParamCount < p_Params.Length)
                        {
                            m_Params[ParamCount] = p_Params[ParamCount];
                        }
                        else
                        {
                            m_Params[ParamCount] = null;
                        }
                    }
                    p_Params = m_Params;

                }
                else if (m_Pars.Length < p_Params.Length)
                {
                    for (Int32 ParamCount = 0; ParamCount < m_Pars.Length; ParamCount++)
                    {
                        m_Params[ParamCount] = p_Params[ParamCount];
                    }
                    p_Params = m_Params;
                }
                m_TransResult = (SysEntity.TransResult)method.Invoke(instance, p_Params);
                if (m_TransResult.isSuccess)
                {
                    m_dtReturn = ((DataTable)m_TransResult.ResultEntity);
                }
            }
            catch (Exception ex)
            {
                throw new ArgumentException("Error " + ex.Message + "! EventAction:" + m_EventEntry.EventAction + ";EventModel:" + m_EventEntry.EventModel, p_QueryMethod);
            }
            return m_dtReturn;
        }

        public SysEntity.TransResult GetDefaultValue(string p_DefaultValue, object[] p_Params, SysEntity.Employee p_Employee)
        {
            SysEntity.TransResult m_TransResult = new SysEntity.TransResult();
            m_TransResult.isSuccess = false;
            SysEntity.EventSetting m_EventEntry = new SysEntity.EventSetting();
            m_EventEntry.EventID = p_DefaultValue;

            try
            {
                SysEntity.TransResult m_TransResultEven = g_BL_System.GetEventSetting(p_Employee, m_EventEntry);
                string m_EventModel = "";
                string m_EventAction = "";
                string m_EventActionType = "";
                string m_EventRef = "";

                if (m_TransResultEven.isSuccess)
                {
                    foreach (DataRow dr in ((DataTable)m_TransResultEven.ResultEntity).Rows)
                    {

                        m_EventEntry.EventModel = dr["EventModel"].ToString();
                        m_EventEntry.EventAction = dr["EventAction"].ToString();

                        m_EventModel = dr["EventModel"].ToString();
                        m_EventAction = dr["EventAction"].ToString();
                        m_EventActionType = dr["EventActionType"].ToString();
                        m_EventRef = dr["EventRef"].ToString();
                        Type m_Type = GetBLType(m_EventModel);
                        object instance = Activator.CreateInstance(m_Type);
                        MethodInfo method = instance.GetType().GetMethod(m_EventAction);
                        m_TransResult = (SysEntity.TransResult)method.Invoke(instance, p_Params);

                    }
                }
               
                
            }
            catch (Exception ex)
            {
                throw new ArgumentException("Error ! EventAction:" + m_EventEntry.EventAction + ";EventModel:" + m_EventEntry.EventModel + "; Message:" + ex.Message, p_DefaultValue);
            }
            return m_TransResult;
        }

        public string GetMultiLanguage(string p_FieldID, string p_Language, SysEntity.Employee p_Employee)
        {
            SysEntity.MultiLanguage p_MultiLanguage = new SysEntity.MultiLanguage();
            p_MultiLanguage.LanguageKey = p_FieldID;
            SysEntity.TransResult m_TransResult = new SysEntity.TransResult();
            string m_LanguageValue = "";
            try
            {
                m_TransResult = g_BL_System.GetMultiLanguage(p_Employee, p_MultiLanguage);
                if (m_TransResult.isSuccess)
                {
                    foreach (DataRow drLan in ((DataTable)m_TransResult.ResultEntity).Rows)
                    {
                        try
                        {
                            m_LanguageValue = drLan["Value" + p_Language].ToString();
                        }
                        catch
                        {
                            m_LanguageValue = drLan["ValueTW"].ToString();
                        }

                    }
                    if (m_LanguageValue == "")
                    {
                        m_LanguageValue = p_FieldID;
                    }
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            return m_LanguageValue;
        }

        public static string GeneratorTR(string p_SectionHtml, string p_SectionType, string p_ControlType, string p_Control, string p_LanguageValue, string p_FieldID, Int32 p_ColumnCount, string p_Width, string p_Height, bool p_isLastItem, Int32 p_ItemIndex, Int32 p_TDIndex, string p_TDClass = "", string p_AuthStyle = "", string p_ValidSource="")
        {
            string m_TDClass = "";
            try
            {
                if (p_TDClass != "")
                {
                    m_TDClass = "class='" + p_TDClass + "'";
                }
                if ((p_ControlType == "TitleTable" || p_ControlType == "TitleGrid" || p_ControlType == "HtmlEditor") || p_SectionType == "Detail" || p_SectionType == "SiteFormStatus")
                {
                    p_Width = "100%";
                }
                string m_TD = " <td>&nbsp;</td><td>&nbsp;</td> ";
                string m_RepairTD = "";
                string m_TDControl = GeneratorTD(p_SectionType, p_ControlType, p_Control, p_LanguageValue, p_FieldID, p_Height, p_ColumnCount, (p_Width == "100%" || p_Width == "0" || p_Width == "display:none"), m_TDClass, p_AuthStyle, p_ValidSource);

                string m_TR_Heigh = "";
                if (p_Height != "")
                {
                    m_TR_Heigh = " Style='heigh: " + p_Height + ";'";
                }

                if ((p_Width != "100%" && p_Width != "0" && p_Width != "display:none"))
                {
                    Int32 m_TDMod = (p_TDIndex % p_ColumnCount);
                    if (m_TDMod == 0)
                    {
                        m_TDControl += " {0} </tr>";
                    }
                    else if (m_TDMod == 1)
                    {
                        m_TDControl = "<tr " + m_TR_Heigh + "> " + m_TDControl + " {0} ";
                    }
                    if (p_isLastItem)
                    {
                        if (m_TDMod == 0)
                        {
                            p_SectionHtml = p_SectionHtml + m_TDControl.Replace("{0}", "");
                        }
                        else
                        {
                            for (Int32 m_Index = 1; (p_ColumnCount - m_TDMod) >= m_Index; m_Index++)
                            {
                                m_RepairTD += m_TD;
                            }
                            if (m_RepairTD != "")
                            {
                                p_SectionHtml = p_SectionHtml + string.Format(m_TDControl, m_RepairTD) + " </tr>";
                            }
                        }
                    }
                    else
                    {
                        p_SectionHtml = p_SectionHtml + string.Format(m_TDControl, "");
                    }
                }
                else
                {
                    if ((p_Width == "0" || p_Width == "display:none"))
                    {
                        m_TDControl = m_TDControl.Insert(m_TDControl.IndexOf("<tr") + 3, " style='display:none;'");
                    }
                    Int32 m_Mod = ((p_TDIndex - 1) % p_ColumnCount);
                    if (m_Mod != 0)
                    {
                        for (Int32 m_Index = 1; (p_ColumnCount - m_Mod) >= m_Index; m_Index++)
                        {
                            m_RepairTD += m_TD;
                        }
                        p_SectionHtml = p_SectionHtml + m_RepairTD + " </tr> " + m_TDControl;
                       
                    }
                    else
                    {
                        if (m_TDControl.IndexOf("{0}") == -1)
                        {
                            p_SectionHtml = p_SectionHtml+ m_TDControl;
                        }
                        else
                        {
                            p_SectionHtml = p_SectionHtml + string.Format(m_TDControl, "");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            return p_SectionHtml;
        }

        public static string GeneratorTD(string p_SectionType, string p_ControlType, string p_Control, string p_LanguageValue, string p_FieldID, string p_Height, Int32 p_ColunmCount, bool p_isTR, string p_TDClass = "", string p_AuthStyle="",string p_ValidSource="")
        {
            string m_HTML = "";
            string m_colspan = "";
            string m_lblCSS = "";
            string m_AllowEmptyCtrl = "";

            try
            {
                if (p_ValidSource == "AllowEmpty")
                {
                    m_lblCSS = p_ValidSource;
                    m_AllowEmptyCtrl = " *";
                }
                string m_TR_Heigh = "";
                if (p_Height != "")
                {
                    m_TR_Heigh = " Style='heigh: " + p_Height + ";'";
                }
                if (p_isTR)
                {
                    m_colspan = " colspan='" + ((p_ColunmCount * 2) - 1).ToString() + "' ";
                }
                if ((p_ControlType == "TitleTable" || p_ControlType == "TitleGrid" || p_ControlType == "HtmlEditor") || p_SectionType == "Detail" || p_SectionType == "SiteFormStatus")
                {
                    m_colspan = " colspan='" + ((p_ColunmCount * 2)).ToString() + "' ";
                    m_HTML = "<td " + m_colspan + "  style='width:100%;text-align:left;" + p_AuthStyle + "'>" + p_Control + "</td>";
                }
                else
                {
                    string m_Br = "";
                    if (p_Control.IndexOf("Remind='Y'")!=-1)
                    {
                        m_Br = " </br> ";
                    }
                    switch (p_SectionType)
                    {
                        case "General":
                            m_HTML = "<td class='lbl' style='width:30%;" + p_AuthStyle + "' " + p_TDClass + ">" + m_Br + "<label class='col-lg-2 control-label " + m_lblCSS + "' id ='lbl" + p_FieldID + "'>" + m_AllowEmptyCtrl + p_LanguageValue + "</label></td><td " + m_colspan + "  style='width:70%;text-align:left;" + p_AuthStyle + "'>" + p_Control + "</td>";
                            break;
                        case "General1":
                            m_HTML = "<td class='lbl' style='width:20%;" + p_AuthStyle + "' " + p_TDClass + ">" + m_Br + "<label class='col-lg-2 control-label " + m_lblCSS + "' id ='lbl" + p_FieldID + "'>" + m_AllowEmptyCtrl + p_LanguageValue + "</label></td><td " + m_colspan + " style='width:30%;text-align:left;" + p_AuthStyle + "'>" + p_Control + "</td>";
                            break;
                        case "General2":
                            m_HTML = "<td class='lbl' style='width:13%;" + p_AuthStyle + "' " + p_TDClass + ">" + m_Br + "<label class='col-lg-2 control-label " + m_lblCSS + "' id ='lbl" + p_FieldID + "'>" + m_AllowEmptyCtrl + p_LanguageValue + "</label></td><td " + m_colspan + "  style='width:20%;text-align:left;" + p_AuthStyle + "'>" + p_Control + "</td>";
                            break;
                    }
                }

                if (p_isTR)
                {
                    m_HTML = "<tr " + m_TR_Heigh + ">" + m_HTML + "</tr>";
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            return m_HTML;
        }

        public SysEntity.TransResult ParserExcel(Stream p_InputStream, string p_fileExt)
        {
            SysEntity.TransResult m_TransResult = new SysEntity.TransResult();
            List<Dictionary<string, object>> m_LisDic = new List<Dictionary<string, object>>();

            try
            {
                ISheet sheet; //Create the ISheet object to read the sheet cell values  
                if (p_fileExt == ".xls")
                {
                    HSSFWorkbook hssfwb = new HSSFWorkbook(p_InputStream); //HSSWorkBook object will read the Excel 97-2000 formats  
                    sheet = hssfwb.GetSheetAt(0); //get first Excel sheet from workbook  
                }
                else
                {
                    XSSFWorkbook hssfwb = new XSSFWorkbook(p_InputStream); //XSSFWorkBook will read 2007 Excel format  
                    sheet = hssfwb.GetSheetAt(0); //get first Excel sheet from workbook   
                }
                if (sheet.LastRowNum != 1 && sheet.LastRowNum != 0)
                {
                    string[] m_ColumnName = new string[sheet.GetRow(0).Cells.Count];
                    for (int row = 0; row <= sheet.LastRowNum; row++) //Loop the records upto filled row  
                    {
                        Dictionary<string, object> m_Dic = new Dictionary<string, object>();
                        if (sheet.GetRow(row) != null) //null is when the row only contains empty cells   
                        {
                            Int32 m_CellIndex = 0;
                            foreach (ICell m_Cell in (sheet.GetRow(row)).Cells)
                            {
                                if (row == 0)
                                {
                                    m_ColumnName[m_CellIndex] = m_Cell.StringCellValue;
                                }
                                else
                                {
                                    m_Dic.Add(m_ColumnName[m_CellIndex], m_Cell.StringCellValue);
                                }
                                m_CellIndex++;
                            }
                            if (row != 0)
                            {
                                m_LisDic.Add(m_Dic);
                            }
                        }
                    }
                }
                m_TransResult.isSuccess = true;
                m_TransResult.ResultEntity = m_LisDic;
            }
            catch (Exception ex)
            {
                throw ex;
            }


            return m_TransResult;
        }

        public T DeserializeObject<T>(string p_Json) where T : new()
        {
            var m_STypet = new T();

            try
            {
                if (p_Json != null)
                {
                    m_STypet = JsonConvert.DeserializeObject<T>(p_Json);
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            return m_STypet;
        }

        public string JsonSerializer(object p_Obj) 
        {
            return JsonConvert.SerializeObject(p_Obj);
        }

        protected void CreateCell(IRow p_IRow, Int32 p_CellIndex, String CellValue, ICellStyle p_CellStyle, Int32 p_Width)
        {
            try
            {
                ICell m_Cell = p_IRow.CreateCell(p_CellIndex);
                m_Cell.SetCellValue(CellValue);
                if (p_CellStyle != null)
                {
                    m_Cell.CellStyle = p_CellStyle;
                }
                if (p_Width != 0)
                {
                    m_Cell.Sheet.SetColumnWidth(p_CellIndex, p_Width);
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }


        public MemoryStream WriteExcelWithNPOI(SysEntity.Employee p_Employee, string p_Language, DataTable p_dt)
        {
            string m_extension = "xls";
            var m_ExportData = new MemoryStream();
            try
            {
                IWorkbook m_Workbook;
                if (m_extension == "xlsx")
                {
                    m_Workbook = new XSSFWorkbook();
                }
                else if (m_extension == "xls")
                {
                    m_Workbook = new HSSFWorkbook();
                }
                else
                {
                    throw new Exception("This format is not supported");
                }

                ISheet m_Sheet1 = m_Workbook.CreateSheet("Sheet 1");
                IRow m_HeadRow = m_Sheet1.CreateRow(0);

                HSSFCellStyle m_HeadCellStyle = (HSSFCellStyle)m_Workbook.CreateCellStyle();
                HSSFFont m_HeadFont= (HSSFFont)m_Workbook.CreateFont();
                m_HeadFont.Boldweight = (short)NPOI.SS.UserModel.FontBoldWeight.Bold;

                m_HeadFont.FontHeightInPoints = 16;
                m_HeadCellStyle.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.Grey25Percent.Index;
                m_HeadCellStyle.FillPattern = NPOI.SS.UserModel.FillPattern.SolidForeground;
                m_HeadCellStyle.SetFont(m_HeadFont);

              
                string[] m_Keys=new string[p_dt.Columns.Count];
                for (int m_ColumnCount = 0; m_ColumnCount < p_dt.Columns.Count; m_ColumnCount++)
                {
                    m_Keys[m_ColumnCount] = p_dt.Columns[m_ColumnCount].ToString();
                }
                Dictionary<string, Object> m_Languages = new Dictionary<string, object>();
                SysEntity.TransResult m_TransResult = g_BL_System.GetListMultiLanguage(p_Employee, p_Language, m_Keys);
                if (m_TransResult.isSuccess )
                {
                    m_Languages = (Dictionary<string, Object>)m_TransResult.ResultEntity;

                }
                Dictionary<Int32, Int32> m_ColumnWidths = new Dictionary<Int32, Int32>();
                Int32 m_BaseWidth = 700;
                for (int m_ColumnCount = 0; m_ColumnCount < p_dt.Columns.Count; m_ColumnCount++)
                {
                    ICell m_HeadCell = m_HeadRow.CreateCell(m_ColumnCount);
                    String m_ColumnName = p_dt.Columns[m_ColumnCount].ToString();
                    string m_Value = m_Languages[m_ColumnName].ToString();
                    m_HeadCell.SetCellValue(m_Value);
                    m_HeadCell.CellStyle = m_HeadCellStyle;
                    m_ColumnWidths[m_ColumnCount] = m_Value.Length * m_BaseWidth;
                }

                HSSFFont m_BodyFont = (HSSFFont)m_Workbook.CreateFont();
                m_BodyFont.FontHeightInPoints = 14;

                NPOI.SS.UserModel.ICellStyle m_CellStyleNORMAL = m_Workbook.CreateCellStyle();
                m_CellStyleNORMAL.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.COLOR_NORMAL;
                m_CellStyleNORMAL.FillPattern = NPOI.SS.UserModel.FillPattern.NoFill;
                m_CellStyleNORMAL.SetFont(m_BodyFont);


                NPOI.SS.UserModel.ICellStyle m_CellStyleLight = m_Workbook.CreateCellStyle();
                m_CellStyleLight.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.LightTurquoise.Index;
                m_CellStyleLight.FillPattern = NPOI.SS.UserModel.FillPattern.SolidForeground;
                m_CellStyleLight.SetFont(m_BodyFont);

                
                for (int m_RowCount = 0; m_RowCount < p_dt.Rows.Count; m_RowCount++)
                {
                    IRow m_BodyRow = m_Sheet1.CreateRow(m_RowCount + 1);
                    for (int m_ColumnCount = 0; m_ColumnCount < p_dt.Columns.Count; m_ColumnCount++)
                    {
                        ICell m_BodyCell = m_BodyRow.CreateCell(m_ColumnCount);
                        String m_ColumnName = p_dt.Columns[m_ColumnCount].ToString();
                        string m_Value = p_dt.Rows[m_RowCount][m_ColumnName].ToString();
                        
                        CheckExcelColumnValue(ref m_Value);
                        m_BodyCell.SetCellValue(m_Value);
                        if (m_Value.Length * m_BaseWidth > m_ColumnWidths[m_ColumnCount])
                        {
                            m_ColumnWidths[m_ColumnCount] = m_Value.Length * m_BaseWidth;
                        }
                        if (m_RowCount % 2 == 0)
                        {
                            m_BodyCell.CellStyle = m_CellStyleNORMAL;
                        }
                        else
                        {
                            m_BodyCell.CellStyle = m_CellStyleLight;
                        }
                    }
                }

                for (int m_ColumnCount = 0; m_ColumnCount < p_dt.Columns.Count; m_ColumnCount++)
                {
                    if (m_ColumnWidths[m_ColumnCount] > 65280)
                    {
                        m_ColumnWidths[m_ColumnCount] = 65280;
                    }
                    m_Sheet1.SetColumnWidth(m_ColumnCount, m_ColumnWidths[m_ColumnCount]);
                }

                m_Workbook.Write(m_ExportData);
            }
            catch(Exception ex)
            {
                m_ExportData.Dispose();
                throw ex;
            }
            
            return m_ExportData;
        }


        public void CheckExcelColumnValue(ref string p_Value)
        {
            byte[] byteStr = Encoding.GetEncoding("big5").GetBytes(p_Value);
            if (p_Value.Length > 250)
            {
                p_Value = p_Value.Substring(0, 250);
                CheckExcelColumnValue(ref p_Value);
            }
            else
            {
                if (Encoding.GetEncoding("big5").GetBytes(p_Value).Length > 250)
                {
                    p_Value = p_Value.Substring(0, p_Value.Length-15);
                    CheckExcelColumnValue(ref p_Value);

                }
            }
        
        }

        public string GetMAIL_SYSTEMADMIN()
        {
            string strAddress = ConfigurationManager.AppSettings["MAIL_SYSTEMADMIN"];
            return strAddress;
        }

        public void SendMail(Business_Logic.Entity.SysEntity.SmtpMail p_SmtpMail)
        {
            try
            {
                // MaillMessage物件....
                string[] strAddress = null;
                MailMessage SmtpMail = new MailMessage();

                // From             
                SmtpMail.From = new MailAddress(ConfigurationManager.AppSettings["MAIL_FROM"]);

                if (!ConfigurationManager.AppSettings["IsProduction"].Equals("Y"))
                {
                    // *********************** Testing Server **************************
                    // TO 
                    SmtpMail.To.Clear();
                    strAddress = ConfigurationManager.AppSettings["MAIL_SYSTEMADMIN"].Split(';');
                    foreach (string i in strAddress)
                    {
                        if (!i.Equals(""))
                        {
                            SmtpMail.To.Add(i.ToString());
                        }
                    }

                    // CC
                    SmtpMail.CC.Clear();
                    // BCC
                    SmtpMail.Bcc.Clear();
                    // Subject
                    SmtpMail.Subject = "[TESTING] " + p_SmtpMail.Subject.Trim();
                    // IsBodyHtml
                    SmtpMail.IsBodyHtml = true;
                    // Priority
                    SmtpMail.Priority = MailPriority.High;
                    // Body
                    SmtpMail.Body = "TO : " + p_SmtpMail.MailTo + "<br>"
                        + "CC : " + p_SmtpMail.MailCC + "<br>"
                        + "<p>" + p_SmtpMail.Body.Trim();
                }
                else
                {
                    SmtpMail.IsBodyHtml = true;
                    // Priority
                    SmtpMail.Priority = MailPriority.High;
                    // ****************** Production Server *************************************
                    //  TO 
                    if (p_SmtpMail.MailTo == "")
                        throw new SystemException("Please Setting TO Mail Address !");
                    strAddress = p_SmtpMail.MailTo.Split(';');
                    foreach (string i in strAddress)
                    {
                        if (!i.Equals(""))
                        {
                            SmtpMail.To.Add(i.ToString());
                        }
                    }

                    // CC                
                    if (p_SmtpMail.MailCC != null && p_SmtpMail.MailCC.Trim() != "")
                    {
                        strAddress = p_SmtpMail.MailCC.Split(';');
                        foreach (string i in strAddress)
                        {
                            if (!i.Equals(""))
                            {
                                SmtpMail.CC.Add(i.ToString());
                            }
                        }
                    }

                    strAddress = ConfigurationManager.AppSettings["MAIL_BCC"].Split(';');
                    foreach (string i in strAddress)
                    {
                        if (!i.Equals(""))
                        {
                            SmtpMail.Bcc.Add(i.ToString());
                        }
                    }


                    // Subject
                    if (p_SmtpMail.Subject == null)
                        throw new SystemException("Please Setting Mail Subject !");
                    SmtpMail.Subject = "[EIP System] " + p_SmtpMail.Subject.Trim();

                    // IsBodyHtml
                    // SmtpMail.IsBodyHtml = p_SmtpMail.isbodyhtml;

                    // Priority
                    SmtpMail.Priority = MailPriority.High;                      //  smptMailTO.priority;

                    // Body
                    if (p_SmtpMail.Body == "")
                        throw new SystemException("Please Setting Mail Contents ! ");
                    SmtpMail.Body = p_SmtpMail.Body.Trim();

                }
                if (p_SmtpMail.Attachment != null)
                {
                    foreach (Dictionary<string, Object> Files in p_SmtpMail.Attachment)
                    {
                        if (Files["FileName"].ToString() != "")
                        {
                            Stream stream = new MemoryStream(Decompress((Byte[])Files["FileEntity"]));
                            Attachment attachment = new Attachment(stream, Files["FileName"].ToString());
                            SmtpMail.Attachments.Add(attachment);
                            attachment.Dispose();
                        }
                    }
                }
                SmtpClient myclient = new SmtpClient();
                myclient.Host = ConfigurationManager.AppSettings["MAIL_HOST"];
                myclient.Port = 25;
                myclient.EnableSsl = false;
                myclient.UseDefaultCredentials = false;
                myclient.DeliveryMethod = SmtpDeliveryMethod.Network;
                myclient.Send(SmtpMail);

            }
            catch
            {
                throw new ArgumentNullException("send mail error");
            }
        }

        public void AddTableColumn(ref DataTable p_dt, Dictionary<string, string> p_DefVal)
        {
            foreach (KeyValuePair<string, string> item in p_DefVal)
            {
                System.Data.DataColumn newColumn = new System.Data.DataColumn(item.Key, typeof(System.String));
                newColumn.DefaultValue = item.Value;
                p_dt.Columns.Add(newColumn);
            }
        }

        public SysEntity.TransResult CheckFormKeys(Dictionary<string, Object> p_DicFormTransfer)
        {
            SysEntity.TransResult m_TransResult = new SysEntity.TransResult();
           
            bool HaveKey = false;
            try
            {
                if (p_DicFormTransfer.ContainsKey("FormKeys"))
                {
                    if ((p_DicFormTransfer["FormKeys"].GetType().BaseType).FullName == "System.Array")
                    {
                        List<Dictionary<string, Object>> m_ListDic = new List<Dictionary<string, object>>();
                        Dictionary<string, Object> m_dic = new Dictionary<string, object>();
                        foreach (object item in (object[])p_DicFormTransfer["FormKeys"])
                        {
                            m_dic = (Dictionary<string, Object>)item;
                            m_ListDic.Add(m_dic);
                        }
                        m_TransResult.isSuccess = true;
                        m_TransResult.ResultEntity = m_ListDic;
                        HaveKey = true;
                    }
                    else
                    {
                        foreach (KeyValuePair<string, Object> item in (Dictionary<string, Object>)p_DicFormTransfer["FormKeys"])
                        {
                            m_TransResult.isSuccess = true;
                            m_TransResult.ResultEntity = p_DicFormTransfer["FormKeys"];
                            HaveKey = true;
                            break;
                        }
                    }
                }
                if (!HaveKey)
                {
                    m_TransResult.isSuccess = false;
                    m_TransResult.LogMessage = "FormKeys is null!!! Please check Table:ViewDetail;Column:ControlRef";
                }
            }
            catch (Exception ex)
            {
                m_TransResult.LogMessage = ex.Message;
                m_TransResult.isSuccess = false;
            }
            return m_TransResult;
        }

        public string GetMaintainHyperlink(Business_Logic.Entity.SysEntity.HyperlinkEntity p_HyperlinkEntity, bool p_IsHOSTNAME=true)
        {
            string m_ResultURL = "";
            string m_KeyNames = "";
            string m_HOSTNAME = ConfigurationManager.AppSettings["HOSTNAME"]; 

            try
            {
                if (p_IsHOSTNAME)
                {
                    m_ResultURL = m_HOSTNAME + @"/KF_Web/GeneratorWebMaintain.aspx?SiteFormID=" + p_HyperlinkEntity.SiteFormID;
                }
                else
                {
                    m_ResultURL = @"/KF_Web/GeneratorWebMaintain.aspx?SiteFormID=" + p_HyperlinkEntity.SiteFormID;
                
                }
                foreach (KeyValuePair<string, string> item in p_HyperlinkEntity.KeyValues)
                {
                    m_KeyNames += (m_KeyNames == "" ? "&KeyNames=" : ",") + item.Key.ToString();
                    m_ResultURL += "&" + item.Key.ToString() + "=" + item.Value.ToString();
                }
                m_ResultURL += m_KeyNames;
            }
            catch 
            {
              
            }
            return m_ResultURL;

        }

        public  byte[] Decompress(byte[] data)
        {
            MemoryStream input = new MemoryStream(data);
            MemoryStream output = new MemoryStream();
            using (GZipStream compressionStream = new GZipStream(input, CompressionMode.Decompress))
            {
                compressionStream.CopyTo(output);
            }
            return output.ToArray();
        }

        public byte[] Compress(byte[] data)
        {
            MemoryStream output = new MemoryStream();
            using (GZipStream compressionStream = new GZipStream(output, CompressionMode.Compress))
            {
                compressionStream.Write(data, 0, data.Length);
            }

            return output.ToArray();
        }

    }
}
