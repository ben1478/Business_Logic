using Business_Logic.Entity;
using NPOI.OpenXmlFormats.Wordprocessing;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Business_Logic.Model
{
    public class BL_Report
    {
        DataComponent g_DC = new DataComponent();





        public SysEntity.TransResult GetCompareAchievement(SysEntity.Employee p_EmployeeEntity, string Year_S, string Year_E, string Type)
        {
            SysEntity.TransResult m_TransResult = new SysEntity.TransResult();
            try
            {
                string m_SQL = "";
                if (Type == "Year")
                {
                    m_SQL = "   select sum( CAST(House_sendcase.get_amount AS DECIMAL))amount, Year(get_amount_date) yyyy     " +
                       "  from House_sendcase    " +
                       "  LEFT JOIN House_apply on House_apply.HA_id = House_sendcase.HA_id AND House_apply.del_tag='0'where House_sendcase.del_tag = '0' " +
                       "  AND isnull(House_sendcase.get_amount,'')<>''  " +
                       "  AND  Year(get_amount_date) between @Year_S and @Year_E ";
                }
                else
                {
                    m_SQL = "   select sum( CAST(House_sendcase.get_amount AS DECIMAL))amount,  " +
                        "    month( get_amount_date)  mm, Year(get_amount_date) yyyy     " +
                        "  from House_sendcase    " +
                        "  LEFT JOIN House_apply on House_apply.HA_id = House_sendcase.HA_id AND House_apply.del_tag='0'where House_sendcase.del_tag = '0' " +
                        "  AND isnull(House_sendcase.get_amount,'')<>''  " +
                        "  AND  Year(get_amount_date) between @Year_S and @Year_E ";
                }

                var m_Params = new
                {
                    Year_S,
                    Year_E
                };
                if (Type == "Year")
                {
                    m_SQL += " group by Year(get_amount_date)   " +
                             " order by Year(get_amount_date) asc   ";
                }
                else
                {
                    m_SQL += " group by Year(get_amount_date), month( get_amount_date) " +
                             " order by Year(get_amount_date) asc, month( get_amount_date)  asc  ";
                }

                using (SqlCommand m_cmd = new SqlCommand(m_SQL))
                {
                    if (Type == "Year")
                    {
                        m_TransResult = g_DC.GetData<ReportEntity.YearAmountInfoByYear>(p_EmployeeEntity, m_cmd, m_Params, "AE_DB");
                        if (m_TransResult.isSuccess)
                        {
                            List<ReportEntity.YearAmountInfoByYear> m_lisResult = (List<ReportEntity.YearAmountInfoByYear>)m_TransResult.ResultEntity;
                            m_TransResult.ResultEntity = m_lisResult;
                        }
                    }
                    else
                    {
                        m_TransResult = g_DC.GetData<ReportEntity.AmountInfo>(p_EmployeeEntity, m_cmd, m_Params, "AE_DB");
                        if (m_TransResult.isSuccess)
                        {
                            List<ReportEntity.AmountInfo> m_lisResult = (List<ReportEntity.AmountInfo>)m_TransResult.ResultEntity;
                            m_TransResult.ResultEntity = m_lisResult;
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



        public SysEntity.TransResult GetAchievementByWorkID(SysEntity.Employee p_EmployeeEntity, string WorkID, string get_amount_date_s, string get_amount_date_e)
        {
            SysEntity.TransResult m_TransResult = new SysEntity.TransResult();
            try
            {
                string m_SQL = " select sum( CAST(House_sendcase.get_amount AS DECIMAL))get_amount ,convert(varchar(7), " +
                    "get_amount_date, 126)  yyyymm , DisplayName from House_sendcase  LEFT JOIN House_apply on House_apply.HA_id " +
                    "= House_sendcase.HA_id AND House_apply.del_tag='0'  LEFT JOIN (select U_num ,U_BC,U_name  FROM User_M where" +
                    " del_tag='0' ) User_M ON User_M.U_num = House_apply.plan_num  where House_sendcase.del_tag = '0'" +
                    " AND isnull(House_sendcase.get_amount,'')<>'' and get_amount_date between @get_amount_date_s and @get_amount_date_E and User_M.U_num=@U_num";
                var m_Params = new
                {
                    get_amount_date_s,
                    get_amount_date_e,
                    U_num = WorkID
                };

                m_SQL += " group by convert(varchar(7), get_amount_date, 126),U_name ";

                using (SqlCommand m_cmd = new SqlCommand(m_SQL))
                {
                    m_TransResult = g_DC.GetData<ReportEntity.Achievement>(p_EmployeeEntity, m_cmd, m_Params, "AE_DB");
                    if (m_TransResult.isSuccess)
                    {
                        List<ReportEntity.Achievement> m_lisAuthMenu = (List<ReportEntity.Achievement>)m_TransResult.ResultEntity;
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

        public SysEntity.TransResult GetAchievementByYear(SysEntity.Employee p_EmployeeEntity, string Year)
        {
            SysEntity.TransResult m_TransResult = new SysEntity.TransResult();
            try
            {
                string get_amount_date_s = Year+"-01-01";
                string get_amount_date_e = Year + "-12-31";

                string m_SQL = " select sum( CAST(House_sendcase.get_amount AS DECIMAL))get_amount ,convert(varchar(7), get_amount_date, 126)  yyyymm  " +
                    " ,convert(varchar(7), get_amount_date, 126)  DisplayName  from House_sendcase " +
                    "   LEFT JOIN House_apply on House_apply.HA_id = House_sendcase.HA_id AND House_apply.del_tag='0'  " +
                    "  where House_sendcase.del_tag = '0' AND isnull(House_sendcase.get_amount,'')<>'' " +
                    "  and get_amount_date between @get_amount_date_s and  @get_amount_date_e  ";
                var m_Params = new
                {
                    get_amount_date_s,
                    get_amount_date_e,
                };

                m_SQL += "  group by  convert(varchar(7), get_amount_date, 126) ";

                using (SqlCommand m_cmd = new SqlCommand(m_SQL))
                {
                    m_TransResult = g_DC.GetData<ReportEntity.Achievement>(p_EmployeeEntity, m_cmd, m_Params, "AE_DB");
                    if (m_TransResult.isSuccess)
                    {
                        List<ReportEntity.Achievement> m_lisAuthMenu = (List<ReportEntity.Achievement>)m_TransResult.ResultEntity;
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


        public SysEntity.TransResult GetAchievementByMonth(SysEntity.Employee p_EmployeeEntity, string Month)
        {
            SysEntity.TransResult m_TransResult = new SysEntity.TransResult();
            try
            {
              

                string m_SQL = " select sum( CAST(House_sendcase.get_amount AS DECIMAL))get_amount ,convert(varchar(7),  " +
                    "get_amount_date, 126)  yyyymm , U_name from House_sendcase  LEFT JOIN House_apply on House_apply.HA_id  " +
                    "= House_sendcase.HA_id AND House_apply.del_tag='0'  LEFT JOIN (select U_num ,U_BC,U_name  FROM User_M where " +
                    " del_tag='0' ) User_M ON User_M.U_num = House_apply.plan_num  where House_sendcase.del_tag = '0'  " +
                    " AND isnull(House_sendcase.get_amount,'')<>''  and convert(varchar(7), get_amount_date, 126) = @Month ";
                var m_Params = new
                {
                    Month
                };
                
                m_SQL += "  group by convert(varchar(7), get_amount_date, 126),U_name,House_apply.plan_num   ";

                using (SqlCommand m_cmd = new SqlCommand(m_SQL))
                {
                    m_TransResult = g_DC.GetData<ReportEntity.Achievement>(p_EmployeeEntity, m_cmd, m_Params, "AE_DB");
                    if (m_TransResult.isSuccess)
                    {
                        List<ReportEntity.Achievement> m_lisAuthMenu = (List<ReportEntity.Achievement>)m_TransResult.ResultEntity;
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


        public SysEntity.TransResult GetAchievementByMonth_BC(SysEntity.Employee p_EmployeeEntity, string Month)
        {
            SysEntity.TransResult m_TransResult = new SysEntity.TransResult();
            try
            {


                string m_SQL = " select sum( CAST(House_sendcase.get_amount AS DECIMAL))get_amount ,convert(varchar(7), get_amount_date, 126)  yyyymm   " +
                    "   ,B.item_D_name DisplayName,User_M.U_BC CheckVal  from House_sendcase   " +
                    "  LEFT JOIN House_apply on House_apply.HA_id = House_sendcase.HA_id AND House_apply.del_tag='0'   " +
                    "   LEFT JOIN (select U_num ,U_BC,U_name  FROM User_M where del_tag='0' ) User_M ON User_M.U_num = House_apply.plan_num    " +
                    "  left join (SELECT [item_D_code],[item_D_name]" +
                    "  FROM [dbo].[Item_list] where [item_M_code]='branch_company' and del_date is null and [item_D_code] <>'') B  ON User_M.U_BC = B.item_D_code " +
                    "  where House_sendcase.del_tag = '0' AND isnull(House_sendcase.get_amount,'')<>''  " +
                    " and convert(varchar(7), get_amount_date, 126) = @Month  ";
                var m_Params = new
                {
                    Month
                };

                m_SQL += "  group by convert(varchar(7), get_amount_date, 126),B.item_D_name,User_M.U_BC order by get_amount desc   ";

                using (SqlCommand m_cmd = new SqlCommand(m_SQL))
                {
                    m_TransResult = g_DC.GetData<ReportEntity.Achievement>(p_EmployeeEntity, m_cmd, m_Params, "AE_DB");
                    if (m_TransResult.isSuccess)
                    {
                        List<ReportEntity.Achievement> m_lisAuthMenu = (List<ReportEntity.Achievement>)m_TransResult.ResultEntity;
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


        public SysEntity.TransResult GetAchievementByYear_BC(SysEntity.Employee p_EmployeeEntity, string Year)
        {
            SysEntity.TransResult m_TransResult = new SysEntity.TransResult();
            try
            {


                string m_SQL = " select sum( CAST(House_sendcase.get_amount AS DECIMAL))get_amount ,convert(varchar(4), get_amount_date, 126)  yyyy   " +
                    "   ,B.item_D_name DisplayName,User_M.U_BC CheckVal  from House_sendcase   " +
                    "  LEFT JOIN House_apply on House_apply.HA_id = House_sendcase.HA_id AND House_apply.del_tag='0'   " +
                    "   LEFT JOIN (select U_num ,U_BC,U_name  FROM User_M where del_tag='0' ) User_M ON User_M.U_num = House_apply.plan_num    " +
                    "  left join (SELECT [item_D_code],[item_D_name]" +
                    "  FROM [dbo].[Item_list] where [item_M_code]='branch_company' and del_date is null and [item_D_code] <>'') B  ON User_M.U_BC = B.item_D_code " +
                    "  where House_sendcase.del_tag = '0' AND isnull(House_sendcase.get_amount,'')<>''  " +
                    " and convert(varchar(4), get_amount_date, 126) = @Year  ";
                var m_Params = new
                {
                    Year
                };

                m_SQL += "  group by convert(varchar(4), get_amount_date, 126),B.item_D_name,User_M.U_BC order by get_amount desc   ";

                using (SqlCommand m_cmd = new SqlCommand(m_SQL))
                {
                    m_TransResult = g_DC.GetData<ReportEntity.Achievement>(p_EmployeeEntity, m_cmd, m_Params, "AE_DB");
                    if (m_TransResult.isSuccess)
                    {
                        List<ReportEntity.Achievement> m_lisAuthMenu = (List<ReportEntity.Achievement>)m_TransResult.ResultEntity;
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



        public SysEntity.TransResult GetAchievementByMonth_BCDtl(SysEntity.Employee p_EmployeeEntity, string Month, string U_BC)
        {
            SysEntity.TransResult m_TransResult = new SysEntity.TransResult();
            try
            {


                string m_SQL = " select sum( CAST(House_sendcase.get_amount AS DECIMAL))get_amount ,convert(varchar(7), get_amount_date, 126)  yyyymm     " +
                    "   ,U_name DisplayName,User_M.U_num CheckVal,U_BC  from House_sendcase     " +
                    "  LEFT JOIN House_apply on House_apply.HA_id = House_sendcase.HA_id AND House_apply.del_tag='0'   " +
                    "  LEFT JOIN (select U_num ,U_BC,U_name  FROM User_M where del_tag='0' ) User_M ON User_M.U_num = House_apply.plan_num      " +
                    "  where House_sendcase.del_tag = '0' AND isnull(House_sendcase.get_amount,'')<>''  " +
                    "  and convert(varchar(7), get_amount_date, 126) = @Month and U_BC=@U_BC ";
                var m_Params = new
                {
                    Month,
                    U_BC
                };

                m_SQL += "   group by convert(varchar(7), get_amount_date, 126),User_M.U_BC,U_name,U_num order by get_amount desc    ";

                using (SqlCommand m_cmd = new SqlCommand(m_SQL))
                {
                    m_TransResult = g_DC.GetData<ReportEntity.Achievement>(p_EmployeeEntity, m_cmd, m_Params, "AE_DB");
                    if (m_TransResult.isSuccess)
                    {
                        List<ReportEntity.Achievement> m_lisAuthMenu = (List<ReportEntity.Achievement>)m_TransResult.ResultEntity;
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


    }
}
