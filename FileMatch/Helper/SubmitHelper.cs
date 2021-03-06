﻿using FileMatch.Entity;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Xml.Linq;
using Utility.Common;
using Utility.Dao;
using Utility.Log;
using System.Data;
using System.Windows.Forms;

namespace Utility.Submit
{
    public class SubmitHelper
    {
        IPEndPoint remotePoint;
        SQLiteDBHelper sqlite;

        public SubmitHelper(SQLiteDBHelper db)
        {
            sqlite = db;
            remotePoint = PublicTool.GetRemoteEp();
        }
        /// <summary>
        /// 纯盘篇提交
        /// </summary>
        /// <param name="code">编号</param>
        /// <param name="dic">标记的参数</param>
        /// <param name="TempDic">临时存储参数,用于提交成功后保存db</param>
        /// <param name="Work_Path">工作路径</param>
        public void ChunPan_Submit(object code,Dictionary<string,object> dic,Dictionary<string,string> TempDic,string Work_Path)
        {
            
            Register reg = new Register();
            string Ini_Path = reg.ReadRegeditKey("QueuePath", "software\\CNKI\\Assistant\\Work\\", Register.RegDomain.CurrentUser).ToString();
            INIManage ini = new INIManage(Ini_Path);
            
            MatchTask task = null;
            foreach (string strValue in ini.SectionValues("Task"))
            {
                string strTask = strValue.Substring(strValue.IndexOf('=') + 1);
                MatchTask mt = strTask.FromJson<MatchTask>();
                if (mt.TaskStatus == "0")
                {
                    task = mt;
                    break;
                }
            }
            TextLog.WritwLog("提交code" + task.Code);

            string filename = Work_Path + "\\Explain.xml";
            string unit = XDocument.Load(filename).Element("ExplainInfo").Element("Info").Value;

            #region 生成提交信息的xml的Content字段

            SendTask sTask = new SendTask();
            sTask.code =task.Code;
            sTask.ArticleCode = code.ToString();
            sTask.Units = unit;
            sTask.Year = dic["学位年度"] == null ? "" : dic["学位年度"].ToString();
            sTask.Level = dic["级别"] == null ? "" : dic["级别"].ToString();
            sTask.IsSecret = dic["保密"].ToString();
            sTask.IsSQ = dic["授权"].ToString();
            sTask.IsQM = dic["签名"].ToString();
            sTask.Iscopyright = dic["版权反馈"].ToString();
            sTask.Explain = dic["备注"] == null ? "" : dic["备注"].ToString();
            sTask.DeleteWords = dic["删除字样"].ToString();
            sTask.DelayDate = dic["滞后上网"].ToString();

            DataRow row = sqlite.ExecuteDataTable("select 小样数,提取页数 from XW_FileOrderinfo where 编号='" + code + "'", null).Rows[0];
            sTask.XiaoYangSum = Convert.ToInt32(row["小样数"]);
            sTask.TotalPage = Convert.ToInt32(row["提取页数"]);

            sTask.Cutf = "否";
            sTask.HardCoverf = "否";

            sTask.TaskComeTime = task.TaskComeTime;
            sTask.IsRead = "";
            sTask.ProcMode = task.ProcMode;
            sTask.SchoolName = "";
            sTask.PaperSummary = "";
            sTask.ProductCode = task.ProductCode;
            sTask.PostName = "文件整理";

            string strSendTask = sTask.ToJson();

            string sql = string.Format("update XW_FileOrderinfo set 年度='{0}',级别='{1}',保密否='{2}',版权反馈否='{3}',是否签名='{4}',是否授权='{5}',备注='{6}',删除字样='{7}',滞后上网='{8}' where 编号='{9}'",
                                       sTask.Year, sTask.Level, sTask.IsSecret, sTask.Iscopyright, sTask.IsQM, sTask.IsSQ, sTask.Explain, sTask.DeleteWords,sTask.DelayDate,code);
            sqlite.ExecuteNonQuery(sql, null);
            TextLog.WritwLog("执行sql成功:"+sql);

            XDocument newdoc = new XDocument();
            XElement node_root = new XElement("ArticleInfo");
            XElement node_code = new XElement("SN");
            XElement node_articleCode = new XElement("ArticleCode");
            XElement node_content = new XElement("CONTENT");
            node_code.Value = code.ToString();
            node_articleCode.Value = code.ToString();
            node_content.Value = strSendTask;
            node_root.Add(node_code,node_articleCode, node_content);
            newdoc.Add(node_root);
            if (!Directory.Exists(Work_Path + "\\ArticlePublish"))
            {
                Directory.CreateDirectory(Work_Path + "\\ArticlePublish");
            }
            newdoc.Save(Work_Path + "\\ArticlePublish\\" + code + ".xml");
            //保存参数,用于在篇提交成功后更新db
            if (TempDic.Keys.Contains(code.ToString()))
                TempDic.Remove(code.ToString());
            TextLog.WritwLog("添加编号:" + code);
            TempDic.Add(code.ToString(), "是");
            string dics = "";
            foreach (string key in TempDic.Keys)
            {
                dics += key + "-";
            }
            TextLog.WritwLog("添加后字典里:" + dics);

            Console.WriteLine(code + ":生成提交信息的xml成功");

            #endregion

            int c0 = Convert.ToInt32(sqlite.ExecuteScalar("select count(*) from db_File where 编号 ='" + code + "'", null));
            string uploadPath = Work_Path + "\\ArticleUpload";
            if (!Directory.Exists(Work_Path + "\\ArticleUpload\\" + code))
            {
                #region 移动篇提交文件至ArticleUpload文件夹
                string tempPath = sqlite.ExecuteScalar("select 路径 from XW_FileOrderinfo where 编号='" + code + "'", null).ToString();

                #region 验证'整理后'文件夹中的文件个数和db中是否一样
                
                int c1 = Traverse(tempPath);
                if (c0 != c1)
                {
                    TextLog.WritwLog(code + ":整理后文件夹中文件个数(" + c1 + ")与db中文件个数(" + c0 + ")不一致");
                    MessageBox.Show(code + ":整理后文件夹中文件个数(" + c1 + ")与db中文件个数(" + c0 + ")不一致,请联系研发!");
                    return;
                }
                #endregion

                DirectoryInfo dir = new DirectoryInfo(tempPath);
               
                if (!Directory.Exists(uploadPath))
                {
                    Directory.CreateDirectory(uploadPath);
                }
                if (!Directory.Exists(uploadPath + "\\" + code))
                {
                    Directory.Move(tempPath, uploadPath + "\\" + code);
                }
                Console.WriteLine(code + ":移动篇提交文件至ArticleUpload文件夹成功");
                #endregion
            }

            #region 验证upload文件夹中的文件个数和db中是否一样
            int c2 = Traverse(uploadPath + "\\" + code);
            if (c0 != c2)
            {
                TextLog.WritwLog(code + ":articleupload文件夹中文件个数("+c2+")与db中文件个数("+c0+")不一致");
                MessageBox.Show(code + ":articleupload文件夹中文件个数(" + c2 + ")与db中文件个数(" + c0 + ")不一致,请联系研发!");
                return;
            }
            #endregion

            #region 给加工助手发送udp进行篇提交
            Console.WriteLine(code + ":开始篇提交");

            //UdpServer udpServer =UdpServer.GetInstance();
            string transport_WorkPath = task.WorkPath.Replace("\\", "\\\\");
            string message = "{\"ArticleCode\":\"" + code + "\"," + "\"Code\":\"" + task.Code + "\"," + "\"LineId\":\"" + task.LineID + "\"," + "\"PostId\":\"" + task.PostID + "\"," +
                "\"TaskComeTime\":\"" + task.TaskComeTime + "\"," + "\"WorkPath\":\"" + transport_WorkPath + "\"}";
            try
            {
                Byte[] sendBytes = Encoding.Default.GetBytes(message);
                PublicTool.localUdp.Send(sendBytes, sendBytes.Length, remotePoint);
                TextLog.WritwLog(message);
                //listView2.Items[code.ToString()].SubItems[1].Text = "提交中";
            }
            catch (Exception ss)
            {
                TextLog.WritwLog(code + "发送udp失败:" + ss.Message);
            }


            #endregion
        }

        public void Kanpan_Submit(object code, Dictionary<string, object> dic, Dictionary<string, string> TempDic, string Work_Path)
        {
            Register reg = new Register();
            string Ini_Path = reg.ReadRegeditKey("QueuePath", "software\\CNKI\\Assistant\\Work\\", Register.RegDomain.CurrentUser).ToString();
            INIManage ini = new INIManage(Ini_Path);

            TextLog.WritwLog("cc1");

            MatchTask task = null;
            foreach (string strValue in ini.SectionValues("Task"))
            {
                string strTask = strValue.Substring(strValue.IndexOf('=') + 1);
                MatchTask mt = strTask.FromJson<MatchTask>();
                if (mt.TaskStatus == "0")
                {
                    task = mt;
                    break;
                }
            }
            TextLog.WritwLog("cc2");

            string filename = Work_Path + "\\Explain.xml";
            string unit = XDocument.Load(filename).Element("ExplainInfo").Element("Info").Value;

            string sql = "update XW_FileOrderinfo set 滞后上网='" + dic["滞后上网"].ToString() + "',保密否='" + dic["保密"].ToString() + "',删除字样='" + dic["删除字样"] + "',论文摘要=\"" + dic["摘要"] + "\" where 编号='" + code + "'";
            //TextLog.WritwLog(sql);
            sqlite.ExecuteNonQuery(sql, null);

            TextLog.WritwLog("cc3");

            #region 生成提交信息的xml的Content字段

            SendTask sTask = new SendTask();
            sTask.code = task.Code;
            sTask.ArticleCode = code.ToString();

            sTask.Units = unit;
            sTask.Year ="";
            sTask.Level = "";
            sTask.IsSecret = dic["保密"].ToString();
            sTask.IsSQ = "";
            sTask.IsQM = "";
            sTask.Iscopyright ="";
            sTask.Explain = "";
            sTask.DeleteWords = dic["删除字样"].ToString();
            sTask.DelayDate = dic["滞后上网"].ToString();

            DataRow row = sqlite.ExecuteDataTable("select 小样数,提取页数 from XW_FileOrderinfo where 编号='" + code + "'", null).Rows[0];
            sTask.XiaoYangSum = Convert.ToInt32(row["小样数"]);
            sTask.TotalPage = Convert.ToInt32(row["提取页数"]);

            sTask.Cutf = "否";
            sTask.HardCoverf = "否";

            sTask.TaskComeTime = task.TaskComeTime;
            sTask.IsRead = "";
            sTask.ProcMode = task.ProcMode;
            sTask.SchoolName = "";
            sTask.PaperSummary = dic["摘要"].ToString();
            sTask.ProductCode = task.ProductCode;
            sTask.PostName = "文件整理";

            string strSendTask = sTask.ToJson();


            XDocument newdoc = new XDocument();
            XElement node_root = new XElement("ArticleInfo");
            XElement node_code = new XElement("SN");
            XElement node_articleCode = new XElement("ArticleCode");
            XElement node_content = new XElement("CONTENT");
            node_code.Value = code.ToString();
            node_articleCode.Value = code.ToString();
            node_content.Value = strSendTask;
            node_root.Add(node_code, node_articleCode, node_content);
            newdoc.Add(node_root);
            if (!Directory.Exists(Work_Path + "\\ArticlePublish"))
            {
                Directory.CreateDirectory(Work_Path + "\\ArticlePublish");
            }
            newdoc.Save(Work_Path + "\\ArticlePublish\\" + code + ".xml");
            //保存参数,用于在篇提交成功后更新db
            if (TempDic.Keys.Contains(code.ToString()))
                TempDic.Remove(code.ToString());
            //TextLog.WritwLog("添加编号:" + code);
            TempDic.Add(code.ToString(), "是");
            string dics = "";
            foreach (string key in TempDic.Keys)
            {
                dics += key + "-";
            }
            //TextLog.WritwLog("添加后字典里:" + dics);

            

            #endregion

            TextLog.WritwLog("cc4");

            int c0 = Convert.ToInt32(sqlite.ExecuteScalar("select count(*) from db_File where 编号 ='" + code + "'", null));
            string uploadPath = Work_Path + "\\ArticleUpload";
            #region 移动篇提交文件至ArticleUpload文件夹
            if (!Directory.Exists(Work_Path + "\\ArticleUpload\\" + code))
            {
                string tempPath = sqlite.ExecuteScalar("select 路径 from XW_FileOrderinfo where 编号='" + code + "'", null).ToString();

                #region 验证'整理后'文件夹中的文件个数和db中是否一样
                
                int c1 = Traverse(tempPath);
                if (c0 != c1)
                {
                    TextLog.WritwLog(code + ":整理后文件夹中文件个数(" + c1 + ")与db中文件个数(" + c0 + ")不一致");
                    MessageBox.Show(code + ":整理后文件夹中文件个数(" + c1 + ")与db中文件个数(" + c0 + ")不一致,请联系研发!");
                    return;
                }
                #endregion

                DirectoryInfo dir = new DirectoryInfo(tempPath);
                
                if (!Directory.Exists(uploadPath))
                {
                    Directory.CreateDirectory(uploadPath);
                }
                if (!Directory.Exists(uploadPath + "\\" + code))
                {
                    Directory.Move(tempPath, uploadPath + "\\" + code);
                }
            }
            #endregion
            TextLog.WritwLog("cc5");

            #region 验证upload文件夹中的文件个数和db中是否一样
            int c2 = Traverse(uploadPath + "\\" + code);
            if (c0 != c2)
            {
                TextLog.WritwLog(code + ":articleupload文件夹中文件个数(" + c2 + ")与db中文件个数(" + c0 + ")不一致");
                MessageBox.Show(code + ":articleupload文件夹中文件个数(" + c2 + ")与db中文件个数(" + c0 + ")不一致,请联系研发!");
                return;
            }
            #endregion

            TextLog.WritwLog("cc6");

            #region 给加工助手发送udp进行篇提交
            

            //UdpServer udpServer =UdpServer.GetInstance();
            string transport_WorkPath = task.WorkPath.Replace("\\", "\\\\");
            string message = "{\"ArticleCode\":\"" + code + "\"," + "\"Code\":\"" + task.Code + "\"," + "\"LineId\":\"" + task.LineID + "\"," + "\"PostId\":\"" + task.PostID + "\"," +
                "\"TaskComeTime\":\"" + task.TaskComeTime + "\"," + "\"WorkPath\":\"" + transport_WorkPath + "\"}";
            try
            {
                Byte[] sendBytes = Encoding.Default.GetBytes(message);
                PublicTool.localUdp.Send(sendBytes, sendBytes.Length, remotePoint);
                TextLog.WritwLog(message);
                //listView2.Items[code.ToString()].SubItems[1].Text = "提交中";
            }
            catch (Exception ss)
            {
                TextLog.WritwLog(code + "发送udp失败:" + ss.Message);
            }


            #endregion
        }

        //public void KanPan_Submit(object code, Dictionary<string, object> dic)
        //{

        //    DataTable dt_text = sqlite.ExecuteDataTable("select 摘要 from db_File where 编号='" + code + "'", null);
        //    TextLog.WritwLog("codechangeSubmit3,code="+code);
        //    string text = "";
        //    foreach (DataRow row in dt_text.Rows)
        //    {
        //        text += row[0].ToString() + " ";
        //    }
        //    text = text.Replace("\"", "");
        //    string sql = "update XW_FileOrderinfo set 保密否='" + dic["保密"].ToString() + "',删除字样='"+dic["删除字样"]+"',论文摘要=\"" + text + "\" where 编号='" + code + "'";

        //    TextLog.WritwLog(sql);
        //    sqlite.ExecuteNonQuery(sql, null);
        //    TextLog.WritwLog("刊盘手动提交第一部");
        //    sql = "update db_State set 保存否='是' where 编号='" + code + "'";
        //    sqlite.ExecuteNonQuery(sql, null);
        //    TextLog.WritwLog("刊盘手动提交第二部");
        //}
        /// <summary>
        /// 纯盘提交不可做任务
        /// </summary>
        /// <param name="code">编号</param>
        /// <param name="info">不可做原因</param>
        /// <param name="Work_Path">工作路径</param>
        public void ChunPan_Submit_UnRead(object code,string info, string Work_Path, Dictionary<string, string> TempDic)
        {
            Register reg = new Register();
            string Ini_Path = reg.ReadRegeditKey("QueuePath", "software\\CNKI\\Assistant\\Work\\", Register.RegDomain.CurrentUser).ToString();
            INIManage ini = new INIManage(Ini_Path);
            string strTask = ini.SectionValues("Task")[0];
            strTask = strTask.Substring(strTask.IndexOf('=') + 1);
            MatchTask task = strTask.FromJson<MatchTask>();
            //var dic = GetBiaoZhu("纯盘");
            string filename = Work_Path + "\\Explain.xml";
            string unit = XDocument.Load(filename).Element("ExplainInfo").Element("Info").Value;

            #region 生成提交信息的xml的Content字段
            SendTask sTask = new SendTask();
            sTask.code = task.Code;
            sTask.ArticleCode = code.ToString();
            sTask.Units = unit;
            
            sTask.IsRead = "否";
            sTask.Year = "";
            sTask.Level = "";
            sTask.IsSecret ="";
            sTask.IsSQ = "";
            sTask.IsQM = "";
            sTask.Iscopyright = "";
            sTask.Explain ="";
            sTask.DeleteWords = "";
            sTask.SchoolName = "";
            sTask.PaperSummary = "";
            sTask.Cutf = "";
            sTask.HardCoverf = "";
            sTask.DelayDate = "";


            sTask.TaskComeTime = task.TaskComeTime;
            //不可做原因说明
            sTask.ProcMode = info;
            
            sTask.ProductCode = task.Code;
            sTask.PostName = "文件整理";

            string strSendTask = sTask.ToJson();

            string sql = string.Format("update XW_FileOrderinfo set 年度='{0}',级别='{1}',保密否='{2}',版权反馈否='{3}',是否签名='{4}',是否授权='{5}',备注='{6}',删除字样='{7}',滞后上网='{8}' where 编号='{9}'",
                                       sTask.Year, sTask.Level, sTask.IsSecret, sTask.Iscopyright, sTask.IsQM, sTask.IsSQ, sTask.Explain, sTask.DeleteWords,sTask.DelayDate, code);
            sqlite.ExecuteNonQuery(sql, null);

            XDocument newdoc = new XDocument();
            XElement node_root = new XElement("ArticleInfo");
            XElement node_code = new XElement("SN");
            XElement node_articleCode = new XElement("ArticleCode");
            XElement node_content = new XElement("CONTENT");
            node_code.Value = code.ToString();
            node_articleCode.Value = code.ToString();
            node_content.Value = strSendTask;
            node_root.Add(node_code, node_content);
            newdoc.Add(node_root);
            if (!Directory.Exists(Work_Path + "\\ArticlePublish"))
            {
                Directory.CreateDirectory(Work_Path + "\\ArticlePublish");
            }
            newdoc.Save(Work_Path + "\\ArticlePublish\\" + code + ".xml");

            //保存参数,用于在篇提交成功后更新db
            if (TempDic.Keys.Contains(code.ToString()))
                TempDic.Remove(code.ToString());
            TempDic.Add(code.ToString(), "否");

            #endregion

            #region 移动篇提交文件至ArticleUpload文件夹
            string tempPath = sqlite.ExecuteScalar("select 路径 from XW_FileOrderinfo where 编号='" + code + "'", null).ToString();

                #region 验证'整理后'文件夹中的文件个数和db中是否一样
            int c0 = Convert.ToInt32(sqlite.ExecuteScalar("select count(*) from db_File where 编号 ='" + code+"'", null));
            int c1 = Traverse(tempPath);
            if (c0 != c1)
            {
                TextLog.WritwLog(code + ":整理后文件夹中文件个数(" + c1 + ")与db中文件个数(" + c0 + ")不一致");
                MessageBox.Show(code + ":整理后文件夹中文件个数(" + c1 + ")与db中文件个数(" + c0 + ")不一致,请联系研发!");
                return;
            }
                #endregion

            DirectoryInfo dir = new DirectoryInfo(tempPath);
            string uploadPath = Work_Path + "\\ArticleUpload";
            if (!Directory.Exists(uploadPath))
            {
                Directory.CreateDirectory(uploadPath);
            }
            if (!Directory.Exists(uploadPath + "\\" + code))
            {
                Directory.Move(tempPath, uploadPath + "\\" + code);
            }
            Console.WriteLine(code + ":移动篇提交文件至ArticleUpload文件夹成功");
            #endregion

            #region 验证upload文件夹中的文件个数和db中是否一样
            int c2 = Traverse(uploadPath + "\\" + code);
            if (c0 != c2)
            {
                TextLog.WritwLog(code + ":articleupload文件夹中文件个数(" + c2 + ")与db中文件个数(" + c0 + ")不一致");
                MessageBox.Show(code + ":articleupload文件夹中文件个数(" + c2 + ")与db中文件个数(" + c0 + ")不一致,请联系研发!");
                return;
            }
            #endregion

            #region 给加工助手发送udp进行篇提交
            Console.WriteLine(code + ":开始篇提交");

            string transport_WorkPath = task.WorkPath.Replace("\\", "\\\\");
            string message = "{\"ArticleCode\":\"" + code + "\"," + "\"Code\":\"" + task.Code + "\"," + "\"LineId\":\"" + task.LineID + "\"," + "\"PostId\":\"" + task.PostID + "\"," +
                "\"TaskComeTime\":\"" + task.TaskComeTime + "\"," + "\"WorkPath\":\"" + transport_WorkPath + "\"}";
            TextLog.WritwLog(code +"要发送的Message是:" + message);
            try
            {
                Byte[] sendBytes = Encoding.Default.GetBytes(message);
                PublicTool.localUdp.Send(sendBytes, sendBytes.Length, remotePoint);
            }
            catch (Exception ss)
            {
                TextLog.WritwLog(code + "发送udp失败:" + ss.Message);
            }


            #endregion
        }

        public void KanPan_Submit_UnRead(object code, string info, string Work_Path, Dictionary<string, string> TempDic)
        {
            //string sql = "update XW_FileOrderinfo set 可读否='否',制作说明=\"" + info+ "\" where 编号='" + code + "'";
            //sqlite.ExecuteNonQuery(sql, null);
            //sql = "update db_State set 保存否='不可做' where 编号='" + code + "'";
            //sqlite.ExecuteNonQuery(sql, null);
            //TextLog.WritwLog("into KanPan_Submit_UnRead");
            Register reg = new Register();
            string Ini_Path = reg.ReadRegeditKey("QueuePath", "software\\CNKI\\Assistant\\Work\\", Register.RegDomain.CurrentUser).ToString();
            INIManage ini = new INIManage(Ini_Path);
            string strTask = ini.SectionValues("Task")[0];
            strTask = strTask.Substring(strTask.IndexOf('=') + 1);
            MatchTask task = strTask.FromJson<MatchTask>();
            //var dic = GetBiaoZhu("纯盘");
            string filename = Work_Path + "\\Explain.xml";
            string unit = XDocument.Load(filename).Element("ExplainInfo").Element("Info").Value;
            //TextLog.WritwLog("基本信息构造完毕");
            #region 生成提交信息的xml的Content字段
            SendTask sTask = new SendTask();
            sTask.code = task.Code;
            sTask.ArticleCode = code.ToString();
            sTask.Units = unit;

            sTask.IsRead = "否";
            sTask.Year = "";
            sTask.Level = "";
            sTask.IsSecret = "";
            sTask.IsSQ = "";
            sTask.IsQM = "";
            sTask.Iscopyright = "";
            sTask.Explain = "";
            sTask.DeleteWords = "";
            sTask.SchoolName = "";
            sTask.PaperSummary = "";
            sTask.Cutf = "";
            sTask.HardCoverf = "";
            sTask.DelayDate = "";


            sTask.TaskComeTime = task.TaskComeTime;
            //不可做原因说明
            sTask.ProcMode = info;

            sTask.ProductCode = task.Code;
            sTask.PostName = "文件整理";

            string strSendTask = sTask.ToJson();

            string sql = string.Format("update XW_FileOrderinfo set 年度='{0}',级别='{1}',保密否='{2}',版权反馈否='{3}',是否签名='{4}',是否授权='{5}',备注='{6}',删除字样='{7}',滞后上网='{8}' where 编号='{9}'",
                                       sTask.Year, sTask.Level, sTask.IsSecret, sTask.Iscopyright, sTask.IsQM, sTask.IsSQ, sTask.Explain, sTask.DeleteWords,sTask.DelayDate, code);
            //TextLog.WritwLog("unread:" + sql);
            sqlite.ExecuteNonQuery(sql, null);
            //TextLog.WritwLog("更新db");
            XDocument newdoc = new XDocument();
            XElement node_root = new XElement("ArticleInfo");
            XElement node_code = new XElement("SN");
            XElement node_articleCode = new XElement("ArticleCode");
            XElement node_content = new XElement("CONTENT");
            node_code.Value = code.ToString();
            node_articleCode.Value = code.ToString();
            node_content.Value = strSendTask;
            node_root.Add(node_code, node_content);
            newdoc.Add(node_root);
            if (!Directory.Exists(Work_Path + "\\ArticlePublish"))
            {
                Directory.CreateDirectory(Work_Path + "\\ArticlePublish");
            }
            newdoc.Save(Work_Path + "\\ArticlePublish\\" + code + ".xml");
            //TextLog.WritwLog("保存xml");
            //保存参数,用于在篇提交成功后更新db
            if (TempDic.Keys.Contains(code.ToString()))
                TempDic.Remove(code.ToString());
            TempDic.Add(code.ToString(), "否");
            //TextLog.WritwLog("同步");

            #endregion

            #region 移动篇提交文件至ArticleUpload文件夹
            string tempPath = sqlite.ExecuteScalar("select 路径 from XW_FileOrderinfo where 编号='" + code + "'", null).ToString();

             #region 验证'整理后'文件夹中的文件个数和db中是否一样
            int c0 = Convert.ToInt32(sqlite.ExecuteScalar("select count(*) from db_File where 编号 ='" + code+"'", null));
            int c1 = Traverse(tempPath);
            if (c0 != c1)
            {
                TextLog.WritwLog(code + ":整理后文件夹中文件个数(" + c1 + ")与db中文件个数(" + c0 + ")不一致");
                MessageBox.Show(code + ":整理后文件夹中文件个数(" + c1 + ")与db中文件个数(" + c0 + ")不一致,请联系研发!");
                return;
            }
                #endregion

            DirectoryInfo dir = new DirectoryInfo(tempPath);
            string uploadPath = Work_Path + "\\ArticleUpload";
            if (!Directory.Exists(uploadPath))
            {
                Directory.CreateDirectory(uploadPath);
            }
            if (!Directory.Exists(uploadPath + "\\" + code))
            {
                Directory.Move(tempPath, uploadPath + "\\" + code);
            }
            TextLog.WritwLog(code + ":移动篇提交文件至ArticleUpload文件夹成功");
            #endregion

            #region 验证upload文件夹中的文件个数和db中是否一样
            int c2 = Traverse(uploadPath + "\\" + code);
            if (c0 != c2)
            {
                TextLog.WritwLog(code + ":articleupload文件夹中文件个数(" + c2 + ")与db中文件个数(" + c0 + ")不一致");
                MessageBox.Show(code + ":articleupload文件夹中文件个数(" + c2 + ")与db中文件个数(" + c0 + ")不一致,请联系研发!");
                return;
            }
            #endregion

            #region 给加工助手发送udp进行篇提交
            //TextLog.WritwLog(code + ":开始篇提交");

            string transport_WorkPath = task.WorkPath.Replace("\\", "\\\\");
            string message = "{\"ArticleCode\":\"" + code + "\"," + "\"Code\":\"" + task.Code + "\"," + "\"LineId\":\"" + task.LineID + "\"," + "\"PostId\":\"" + task.PostID + "\"," +
                "\"TaskComeTime\":\"" + task.TaskComeTime + "\"," + "\"WorkPath\":\"" + transport_WorkPath + "\"}";
            //TextLog.WritwLog(message);
            try
            {
                Byte[] sendBytes = Encoding.Default.GetBytes(message);
                PublicTool.localUdp.Send(sendBytes, sendBytes.Length, remotePoint);
            }
            catch (Exception ss)
            {
                TextLog.WritwLog(code + "发送udp失败:" + ss.Message);
            }


            #endregion
        }

        public int Traverse(string sPathName)
        {
            int i = 0;
            //创建一个队列用于保存子目录
            Queue<string> pathQueue = new Queue<string>();
            pathQueue.Enqueue(sPathName);
            //开始循环查找文件，直到队列中无任何子目录
            while (pathQueue.Count > 0)
            {
                DirectoryInfo diParent = new DirectoryInfo(pathQueue.Dequeue());
                foreach (DirectoryInfo diChild in diParent.GetDirectories())
                    pathQueue.Enqueue(diChild.FullName);
                foreach (FileInfo fi in diParent.GetFiles())
                {
                    if(fi.Extension!=".ini"&&fi.Extension!=".xml")
                    i++;
                }
            }
            return i;
        }
    }
}
