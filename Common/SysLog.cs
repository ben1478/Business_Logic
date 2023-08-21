using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Xml;
using System.Data;
using System.Diagnostics;
using Business_Logic.Entity;

namespace Business_Logic.Common
{


    public class SysLog
    {
        /// <summary>
        /// Log存放位置
        /// </summary>
        public static string DBCONFIG_FILEPATH = @"C:\EIP_LOG\";
        /// <summary>
        /// 寫入LOG
        /// </summary>
        /// <param name="dtInfo">錯誤訊息LogEntity</param>
        public void EventLogManager(SysEntity.EventLog p_LogEntity)
        {

            string DirPath = DBCONFIG_FILEPATH + "\\" + DateTime.Now.ToString("yyyyMMdd") + "\\" + p_LogEntity.Employee.CompanyCode + "\\" + p_LogEntity.Employee.WorkID;
            string EventFilePath = DirPath + "\\Log_" + DateTime.Now.ToString("yyyMMddHH") + ".xml";
            Directory.CreateDirectory(DirPath);

            XmlDocument xmlDoc = new XmlDocument();
            try
            {
                //判斷是否檔案存在
                if (!File.Exists(EventFilePath))
                {
                    xmlDoc.LoadXml("<?xml version=\"1.0\"?><Log></Log>");
                }
                else
                {
                    //針對已存在XML檔案寫入錯誤相關資訊
                    FileStream fs = null;
                    try
                    {
                        fs = new FileStream(EventFilePath, FileMode.OpenOrCreate, FileAccess.ReadWrite, FileShare.ReadWrite);
                        xmlDoc.Load(fs);
                    }
                    catch
                    {
                        //當XML檔案格式錯誤時重新建立Root
                        xmlDoc.LoadXml("<?xml version=\"1.0\"?><Log></Log>");
                    }
                    finally
                    {
                        fs.Close();
                    }
                }
                XmlNode note = xmlDoc.SelectSingleNode("Log");
                XmlElement xn = xmlDoc.CreateElement("Event");
                // xn.SetAttribute("runStartTime", this._runStartDateTime.ToString());
                xn.SetAttribute("runEndTime", DateTime.Now.ToString());

                XmlElement xr = xmlDoc.CreateElement("runExecuteInfo");
                XmlElement sub = xmlDoc.CreateElement("LogID");
                sub.InnerText = p_LogEntity.LogResult;
                xr.AppendChild(sub);
                XmlElement sub1 = xmlDoc.CreateElement("TransResult");
                sub1.InnerText = p_LogEntity.LogResult;
                xr.AppendChild(sub1);
                XmlElement sub2 = xmlDoc.CreateElement("LogMessage");
                sub2.InnerText = p_LogEntity.LogMessage;
                xr.AppendChild(sub2);
               
                XmlElement sub3 = xmlDoc.CreateElement("Sql");
                sub3.InnerText = p_LogEntity.Sql;
                xr.AppendChild(sub3);
                xn.AppendChild(xr);
                note.AppendChild(xn);

                xmlDoc.Save(EventFilePath);
            }
            catch (Exception ex)
            {
                EventLog.WriteEntry("EIP", "Exception Message:" + ex.Message + "; Exception StackTrace:" + ex.StackTrace, EventLogEntryType.Information);
            }
        }
    }
}
