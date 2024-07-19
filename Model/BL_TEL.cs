using Business_Logic.Entity;
using NPOI.OpenXmlFormats.Dml.Diagram;
using NPOI.SS.Formula.Functions;
using NPOI.XWPF.UserModel;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;

namespace Business_Logic.Model
{
    public class BL_TEL
    {
        DataComponent g_DC = new DataComponent();


        public class TelData_M
        {
            public String TM_id { get; set; }
            public String CS_Name { get; set; }
            public String CS_Remark { get; set; }
            public String CS_ID { get; set; }
            public String CS_PID { get; set; }
            public String CS_MTEL1 { get; set; }
            public String CS_MTEL2 { get; set; }
            public String CS_Register_Address { get; set; }
            public List< TelData_D> TelData_D { get; set; }
        }

        public class TelData_D
        {
            public String LogKey { get; set; }
            public String TM_id { get; set; }
            public String TM_D_id { get; set; }
            public String Call_type { get; set; }
            public String Fin_type { get; set; }
            public String Memo { get; set; }
            public String U_name { get; set; }
            public String Call_date_S { get; set; }
            public String Call_date_E { get; set; }
            public String Call_name { get; set; }
            public String Fin_name { get; set; }
          

        }


        public SysEntity.TransResult GetCall_CenterGrid(SysEntity.Employee p_EmployeeEntity, Dictionary<string, Object> p_dicRec, string p_PageIndex, string p_PageSize, Dictionary<string, Object> p_Param)
        {
            SysEntity.TransResult m_TransResult = new SysEntity.TransResult();
            Dictionary<string, Object> m_objBetween = new Dictionary<string, Object>();
           
            string[] m_LikeProp = { "CS_name" };
            ArrayList m_alGrid = new ArrayList();
            try
            {
                string m_SQL = " SELECT tm.TM_type, tm.ha_id,td.TM_id,td.TM_D_id, ha.CS_name,ha.cs_id,ha.CS_MTEL1 ,New_HA_id,CS_remark,InDate,cs_register_address,isnull(NewHA.show_plan_type, '')show_plan_type,                                     " +
                                "        convert(varchar(4),(convert(varchar(4), Tel_Assign_date, 126)-1911))+'/'+convert(varchar(2), month(Tel_Assign_date)) +'/'+convert(varchar(2), day(Tel_Assign_date)) Tel_Assign_date_DIS ,      " +
                                "        L._count ,il.item_D_name AS call_name,il_1.item_D_name AS fin_name,ISNULL(td.save_day, 0) save_day,                           " +
                                "        ISNULL(pre_count, 0)pre_count /*估價狀態*/ ,                                                                                                                                                   " +
                                "        CASE WHEN ISNULL(pre_count, 0)=0 THEN '0'                                                                                                                                                      " +
                                "             WHEN ISNULL(pre_count, 0)=1 THEN /*只有一筆直接顯示狀態*/                                                                                                                                 " +
                                "                   (SELECT (SELECT item_D_name FROM Item_list WHERE item_M_code = 'pre_process_type' AND item_D_type='Y'                                                                               " +
                                "                         AND item_D_code = pre_process_type AND show_tag='0' AND del_tag='0') AS show_pre_process_type                                                                                 " +
                                "                    FROM House_pre WHERE del_tag = '0' AND HA_id = New_HA_id )                                                                                                                         " +
                                "            ELSE /*一筆以上另外處理*/ 'S'                                                                                                                                                              " +
                                "        END pre_status ,                                                                                                                                                                               " +
                                "        CASE WHEN ISNULL(pre_count, 0)=1 THEN                                                                                                                                                          " +
                                "            (SELECT HP_id FROM House_pre WHERE del_tag = '0' AND HA_id= New_HA_id)                                                                                                                     " +
                                "        END HP_id ,                                                                                                                                                                                    " +
                                "        CASE WHEN ISNULL(pre_count, 0)=1 THEN                                                                                                                                                          " +
                                "             (SELECT HP_cknum FROM House_pre WHERE del_tag = '0' AND HA_id= New_HA_id)                                                                                                                 " +
                                "        END HP_cknum                                                                                                                                                                                   " +
                                " FROM Telemarketing_M tm                                                                                                                                                                               " +
                                " JOIN view_Telemarketing_Curr td ON tm.TM_id = td.tm_id                                                                                                                                                " +
                                " JOIN view_Telemarketing_source ha ON tm.HA_id = ha.HA_id                                                                                                                                              " +
                                " AND tm.TM_type=ha.TM_type                                                                                                                                                                             " +
                                " LEFT JOIN Item_list il ON td.Call_type = il.item_D_code                                                                                                                                               " +
                                " AND il.item_M_code = 'call_type'                                                                                                                                                                      " +
                                " LEFT JOIN Item_list il_1 ON td.Fin_type = il_1.item_D_code                                                                                                                                            " +
                                " AND il_1.item_M_code = 'fin_type'                                                                                                                                                                     " +
                                " LEFT JOIN Item_list il_2 ON TelSour = il_2.item_D_code                                                                                                                                                " +
                                " AND il_2.item_M_code = 'TelSour'                                                                                                                                                                      " +
                                " LEFT JOIN Item_list il_3 ON TelAsk = il_3.item_D_code                                                                                                                                                 " +
                                " AND il_3.item_M_code = 'TelAsk'                                                                                                                                                                       " +
                                " LEFT JOIN                                                                                                                                                                                             " +
                                "   (SELECT TM_id,count(TM_D_id)_count FROM Telemarketing_log t_log GROUP BY TM_id) L ON Tm.TM_id=L.TM_id                                                                                               " +
                                " LEFT JOIN                                                                                                                                                                                             " +
                                "   (SELECT house_apply.TM_id,house_apply.HA_id New_HA_id,HA_cknum,                                                                                                                                     " +
                                "      (SELECT item_D_name FROM Item_list WHERE item_M_code = 'plan_type' AND item_D_type='Y'                                                                                                           " +
                                "         AND item_D_code = House_apply.plan_type AND show_tag='0' AND del_tag='0') AS show_plan_type,                                                                                                  " +
                                " 	(SELECT count(*) FROM House_pre WHERE del_tag='0' AND HA_id = House_apply.HA_id) AS HP_num                                                                                                          " +
                                " 		FROM House_apply WHERE House_apply.del_tag = '0' AND House_apply.plan_num = 'K0260') NewHA ON tm.TM_id=NewHA.TM_id                                                                              " +
                                " LEFT JOIN                                                                                                                                                                                             " +
                                "   (SELECT HA_id,count(HA_id)pre_count FROM House_pre WHERE del_tag = '0' GROUP BY HA_id) PP ON New_HA_id=PP.HA_id                                                                                     " +
                                " WHERE tm.Assign_num = 'K0260' AND td.Call_num = 'K0260' AND isnull(isRec, '') =''  ";

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

        public DataTable GetTelemarketing_source(string TM_type, string HA_id)
        {
            DataTable m_RtnDT = new DataTable();
            try
            {
                string m_SQL = " select TM_id, cs_name,CS_remark, CS_ID CS_PID, CS_MTEL1, cs_mtel2, CS_register_address  from view_Telemarketing_source T " +
                               "   left join　Telemarketing_M　M on T.ha_id=M.ha_id and T.TM_type=M.TM_type where t.TM_type=@TM_type and t.HA_id = @HA_id  ";
                using (SqlCommand m_cmd = new SqlCommand(m_SQL))
                {
                    m_cmd.Parameters.Add(new SqlParameter("TM_type", TM_type));
                    m_cmd.Parameters.Add(new SqlParameter("HA_id", HA_id));
                    m_RtnDT = g_DC.GetDataTable(m_cmd, "AE_DB");
                }
            }
            catch
            {
                throw;
            }
            return m_RtnDT;
        }

        public SysEntity.TransResult get_csName_detail(SysEntity.Employee p_EmployeeEntity, string TM_type, string HA_id)
        {
            SysEntity.TransResult m_TransResult = new SysEntity.TransResult();
            try
            {
                TelData_M m_TelData_M = new TelData_M();
                string TM_id = "";
                DataTable m_TSResult = GetTelemarketing_source(TM_type, HA_id);

                foreach (DataRow dr in m_TSResult.Rows)
                {
                    TM_id = dr["TM_id"].ToString();
                    m_TelData_M.TM_id = dr["TM_id"].ToString();
                    m_TelData_M.CS_Name = dr["CS_Name"].ToString();
                    m_TelData_M.CS_MTEL1 = dr["CS_MTEL1"].ToString();
                    m_TelData_M.CS_MTEL2 = dr["CS_MTEL2"].ToString();
                    m_TelData_M.CS_Remark = dr["CS_Remark"].ToString();
                    m_TelData_M.CS_ID = dr["CS_PID"].ToString();
                    m_TelData_M.CS_Register_Address = dr["CS_Register_Address"].ToString();
                }

                var m_Params = new
                {
                    TM_id
                };

                string m_SQL = "  select tl.LogKey, tl.TM_id, tl.TM_D_id, tl.Call_type, tl.Fin_type, tl.Memo, um.U_name   ,  " +
                        "  convert(varchar, tl.Call_date_S, 120) Call_date_S , convert(varchar, tl.Call_date_E, 120)  " +
                        "  Call_date_E  , il.item_D_name as call_name, il_1.item_D_name as fin_name  from Telemarketing_Log tl  " +
                        "  join User_M um on tl.Call_num = um.U_num  join Item_list il on tl.Call_type = il.item_D_code and il.item_M_code = 'call_type'  " +
                        "   join Item_list il_1 on tl.Fin_type = il_1.item_D_code and il_1.item_M_code = 'fin_type'  where TM_id = @TM_id order by call_date_s desc   ";
                       

                using (SqlCommand m_cmd = new SqlCommand(m_SQL))
                {
                    m_TransResult = g_DC.GetData<TelData_D>(p_EmployeeEntity, m_cmd, m_Params, "AE_DB");
                    if (m_TransResult.isSuccess)
                    {
                        List<TelData_D> m_lisAuthMenu = (List<TelData_D>)m_TransResult.ResultEntity;
                        m_TelData_M.TelData_D = m_lisAuthMenu;
                        m_TransResult.ResultEntity = m_TelData_M;
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
