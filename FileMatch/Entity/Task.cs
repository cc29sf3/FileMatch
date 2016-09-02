using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace FileMatch.Entity
{
    /// <summary>
    /// 取得的任务类
    /// </summary>
    public class MatchTask
    {
       public string ArticleCode { get; set; }
       public string Code { get; set; }
       public string DownTime { get; set; }
       public string FinishTime { get; set; }
       public string KeyName { get; set; }
       public string LineID { get; set; }
       public string PostID { get; set; }
       public string PostName { get; set; }
       public string ProcMode { get; set; }
       public string ProcType { get; set; }
       public string ProductCode { get; set; }
       public string Reason { get; set; }
       public string StartTime { get; set; }
       public string TaskCode { get; set; }
       public string TaskComeTime { get; set; }
       public string TaskKind { get; set; }
       public string TaskLevel { get; set; }
       public string TaskStatus { get; set; }
       public string UpPath { get; set; }
       public string WorkPath { get; set; }
    }
    /// <summary>
    /// 生成的任务类
    /// </summary>
    public class SendTask
    { 
        //public long code{ get; set; }//编号
        public string code { get; set; }//编号
        //public long ProductCode { get; set; }
        public string ProductCode { get; set; }
        public string PostName { get; set; }
        //public long ArticleCode { get; set; }//接收编号
        public string ArticleCode { get; set; }//接收编号
        public string Units { get; set; }//授予单位
        public string Year { get; set; }//年度
        public string Level { get; set; }//级别
        public string IsSecret { get; set; }//是否保密
        public string Iscopyright { get; set; }//版权反馈否 
        public string IsSQ { get; set; }//是否授权
        public string IsQM { get; set; }//有无签名
        public string TaskComeTime { get; set; }//任务到岗时间
        public string Explain { get; set; }//备注
        public string IsRead { get; set; }
        public string ProcMode { get; set; }//制作说明
        public string SchoolName { get; set; }//学院名称
        public string PaperSummary { get; set; }//摘要
        public string DeleteWords { get; set; }//删除字样
        public string Cutf{get;set;}//  可拆切
        public string HardCoverf{get;set;}  //精装
        public int XiaoYangSum { get; set; }//小样数
        public int TotalPage { get; set; }//提取页数
        public string DelayDate { get; set; }//滞后上网
    }

    public class SubmitResult
    {
        public string ArticleCode{get;set;}
        public string Statu{get;set;}
        public string ErrInfo{get;set;}
    }
}
