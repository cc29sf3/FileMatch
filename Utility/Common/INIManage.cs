using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;

namespace Utility.Common
{
    /// <summary>
    /// 描述：Ini文件的操作
    /// 作者：员战强
    /// 日期：2014-06-17
    /// </summary>
    public class INIManage
    {
        private string m_Path = null;		//ini文件路径
        public INIManage(string iniPath)
        {
            this.m_Path = iniPath;
        }
        #region 段信息的获取
        //读取一个ini文件中的所有段
        [DllImport("kernel32", EntryPoint = "GetPrivateProfileSectionNamesW", CharSet = CharSet.Unicode)]
        private extern static int getSectionNames(
        [MarshalAs(UnmanagedType.LPWStr)] string szBuffer, int nlen, string filename);
        //读取段里的所有数据
        [DllImport("kernel32", EntryPoint = "GetPrivateProfileSectionW", CharSet = CharSet.Unicode)]
        private extern static int getSectionValues(string section,
        [MarshalAs(UnmanagedType.LPWStr)] string szBuffer, int nlen, string filename);
        #endregion

        #region 键值的获取和设置
        //读取键的整形值
        [DllImport("kernel32", EntryPoint = "GetPrivateProfileIntW", CharSet = CharSet.Unicode)]
        private static extern int getKeyIntValue(string section, string key, int nDefault, string fileName);
        //读取字符串键值
        [DllImport("kernel32", EntryPoint = "GetPrivateProfileStringW", CharSet = CharSet.Unicode)]
        private extern static int getKeyValue(string section, string key, int lpDefault,
            [MarshalAs(UnmanagedType.LPWStr)] string szValue, int nlen, string filename);
        //写字符串键值
        [DllImport("kernel32", EntryPoint = "WritePrivateProfileStringW", CharSet = CharSet.Unicode)]
        private static extern bool setKeyValue(string section, string key, string szValue, string fileName);
        //写段值
        [DllImport("kernel32", EntryPoint = "WritePrivateProfileSectionW", CharSet = CharSet.Unicode)]
        private static extern bool setSectionValue(string section, string szvalue, string filename);
        /// <summary>
        /// 加密
        /// </summary>
        /// <param name="szInput"></param>
        /// <param name="szOutput"></param>
        /// <returns></returns>
        [DllImport("UnicodeBase64", EntryPoint = "StringEncodeBase64", CharSet = CharSet.Unicode)]
        private static extern int StringEncodeBase64(string szInput, StringBuilder szOutput);
        /// <summary>
        /// 解密
        /// </summary>
        /// <param name="szInput"></param>
        /// <param name="szOutput"></param>
        /// <returns></returns>
        [DllImport("UnicodeBase64", EntryPoint = "StringDecodeBase64", CharSet = CharSet.Unicode)]
        private static extern int StringDecodeBase64(string szInput, StringBuilder szOutput);
        #endregion
        /// <summary>
        /// Ini加密
        /// </summary>
        /// <param name="key">需加密字符</param>
        /// <returns></returns>
        public string IniEncode(string key)
        {
            StringBuilder buffer = new StringBuilder(102400);
            var result = string.Empty;
            try
            {
                var a = StringEncodeBase64(key, buffer);
                result = buffer.ToString();
            }
            catch (Exception ex)
            {
                var b = ex.Message;
                result = "";
            }
            return result;
        }
        /// <summary>
        /// Ini解密
        /// </summary>
        /// <param name="key">需解密字符</param>
        /// <returns></returns>
        public string IniDecode(string key)
        {
            StringBuilder buffer = new StringBuilder(102400);
            var result = string.Empty;
            try
            {
                var a = StringDecodeBase64(key, buffer);
                result = buffer.ToString();
            }
            catch (Exception ex)
            {
                result = "";
            }
            return result;
        }
        private static readonly char[] sept = { '\0' };	//分隔字符
        /// <summary>
        /// 读取所有段名
        /// </summary>
        public string[] SectionNames
        {
            get
            {
                string buffer = new string('\0', 32768);
                int nlen = getSectionNames(buffer, 32768 - 1, m_Path) - 1;
                if (nlen > 0)
                {
                    return buffer.Substring(0, nlen).Split(sept);
                }
                return null;
            }
        }
        /// <summary>
        /// 读取段里的数据到一个字符串数组
        /// </summary>
        /// <param name="section">段名</param>
        /// <param name="bufferSize">读取的数据大小(字节)</param>
        /// <returns>成功则不为null</returns>
        public string[] SectionValues(string section, int bufferSize)
        {
            string buffer = new string('\0', bufferSize);
            int nlen = getSectionValues(section, buffer, bufferSize, m_Path) - 1;
            if (nlen > 0)
            {
                return buffer.Substring(0, nlen).Split(sept);
            }
            return null;
        }
        public string[] SectionValues(string section)
        {
            return SectionValues(section, 32768);
        }

        /// <summary>
        /// 从一个段中读取其 键-值 数据
        /// </summary>
        /// <param name="section">段名</param>
        /// <param name="bufferSize">读取的数据大小(字节)</param>
        /// <returns>成功则不为null</returns>
        public Dictionary<string, string> SectionValuesEx(string section, int bufferSize)
        {
            string[] sztmp = SectionValues(section, bufferSize);
            if (sztmp != null)
            {
                int ArrayLen = sztmp.Length;
                if (ArrayLen > 0)
                {
                    Dictionary<string, string> dtRet = new Dictionary<string, string>();
                    for (int i = 0; i < ArrayLen; i++)
                    {
                        var splitArray = sztmp[i].Split('=');
                        if (splitArray.Count() == 2)
                        {
                            dtRet.Add(splitArray[0], splitArray[1]);
                        }
                    }
                    return dtRet;
                }
            }
            return new Dictionary<string, string>();
        }
        public Dictionary<string, string> SectionValuesEx(string section)
        {
            bool isCanWrite = false;
            while (!isCanWrite)
            {
                isCanWrite = this.m_Path.IsCanWrite();
            }
            return SectionValuesEx(section, 1024000);
        }

        /// <summary>
        /// 写一个段的数据
        /// </summary>
        /// <param name="section"></param>
        /// <param name="szValue">段的数据(如果为null则删除这个段)</param>
        /// <returns>成功则为true</returns>
        public bool setSectionValue(string section, string szValue)
        {
            return setSectionValue(section, szValue, m_Path);
        }

        /// <summary>
        /// 读整形键值
        /// </summary>
        /// <param name="section"></param>
        /// <param name="key"></param>
        /// <returns>成功则不为-1</returns>
        public int getKeyIntValue(string section, string key)
        {
            return getKeyIntValue(section, key, -1, m_Path);
        }

        /// <summary>
        /// 写整形键值
        /// </summary>
        /// <param name="section"></param>
        /// <param name="key"></param>
        /// <param name="dwValue"></param>
        /// <returns>成功则为true</returns>
        public bool setKeyIntValue(string section, string key, int dwValue)
        {
            return setKeyValue(section, key, dwValue.ToString(), m_Path);
        }

        /// <summary>
        /// 读取键值
        /// </summary>
        /// <param name="section"></param>
        /// <param name="key"></param>
        /// <returns>成功则不为null</returns>
        public string getKeyValue(string section, string key)
        {
            string szBuffer = new string('0', 1024);
            int nlen = getKeyValue(section, key, 0, szBuffer, 1100, m_Path);
            return szBuffer.Substring(0, nlen);
        }

        /// <summary>
        /// 写字符串键值
        /// </summary>
        /// <param name="section"></param>
        /// <param name="key"></param>
        /// <param name="szValue"></param>
        /// <returns>成功则为true</returns>
        public bool setKeyValue(string section, string key, string szValue)
        {
            bool isCanWrite = false;
            while (!isCanWrite)
            {
                isCanWrite = this.m_Path.IsCanWrite();
            }
            return setKeyValue(section, key, szValue, m_Path);
        }
    }
}
