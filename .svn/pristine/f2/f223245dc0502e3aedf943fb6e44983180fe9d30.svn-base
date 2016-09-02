using System;
using System.IO;
using System.Text;

namespace Utility.Log
{
    /// <summary>
    /// 描述：记录本地日志
    /// 作者：员战强
    /// 日期：2014-03-14
    /// </summary>
    public static class TextLog
    {
        private static object lockObject = new object();
        private static string Msgs;

        /// <summary>
        /// 添加日志
        /// </summary>
        /// <param name="msgs">日志内容</param>
        /// <param name="isError">是否错误</param>
        /// <remarks>
        /// 2014-02-28 add by yzq
        /// </remarks>
        public static void WritwLog(string msgs, bool isError = false)
        {
            if (Msgs != msgs && msgs != "" && msgs != null)
            {
                Add(isError, msgs);
                Msgs = msgs;
            }
        }

        /// <summary>
        /// 添加日志
        /// </summary>
        /// <param name="isError">是否错误</param>
        /// <param name="msgs">日志内容</param>
        /// <remarks>
        /// 2014-02-28 add by yzq
        /// </remarks>
        private static void Add(bool isError, params object[] msgs)
        {
            string savePath = AppDomain.CurrentDomain.BaseDirectory + "\\log\\";
            if (!Directory.Exists(savePath))
            {
                Directory.CreateDirectory(savePath);
            }
            string fn = savePath + DateTime.Now.ToString("yyyy-MM-dd") + ".log";
            if (msgs == null || msgs.Length == 0)
                return;
            StringBuilder sb = new StringBuilder();
            for (int i = 0; i < msgs.Length; i++)
                sb.AppendFormat("{{{0}}}\r\n", i);
            sb.Append("\r\n");
            var time = DateTime.Now;
            string messageFormat = string.Empty;
            messageFormat += "[" + time.Hour + ":" + time.Minute + ":" + time.Second + ":" + time.Millisecond + "]";
            if (isError)
                messageFormat += "Error:";
            else
                messageFormat += "Message:";
            messageFormat += sb.ToString();
            lock (lockObject)
                File.AppendAllText(fn, string.Format(messageFormat, msgs));
        }
      
    }
}