using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
//using Assistant.Utility.Config;
using Microsoft.Win32;

namespace Utility.Common
{
    /// <summary>
    /// 描述：注册表操作类
    ///       在使用的过程中，主要用到键值的：判断、写、读 就可以
    /// 作者：员战强
    /// 日期：2014-02-27
    /// </summary>
    public class Register
    {
        #region 操作方法

        #region 注册项操作
        /// <summary>
        /// 创建注册表项
        /// 例子：如regDomain是HKEY_CLASSES_ROOT，
        ///       subkey是software\\CNKI\\，
        ///       则将创建HKEY_CLASSES_ROOT\\software\\CNKI\\注册表项 
        /// </summary> 
        /// <param name="subKey">注册表项名称（例：software\\CNKI\\）</param>
        /// <param name="regDomain">注册表基项域</param> 
        /// <remarks>
        /// 2014-02-27 add by yzq
        /// </remarks>
        public void CreateSubKey(string subKey, RegDomain regDomain)
        {
            //判断注册表项名称是否为空，如果为空，返回false
            if (string.IsNullOrEmpty(subKey))
            {
                return;
            }
            //创建基于注册表基项的节点 
            RegistryKey key = GetRegDomain(regDomain);
            //要创建的注册表项的节点 
            RegistryKey sKey = null;
            if (!IsSubKeyExist(subKey, regDomain))
            {
                sKey = key.CreateSubKey(subKey);
            }
            if (sKey != null)
            {
                sKey.Close();
            }
            //关闭对注册表项的更改
            key.Close();
        }
        /// <summary>
        /// 判断注册表项是否存在
        /// 例子：如regDomain是HKEY_CLASSES_ROOT，
        ///       subkey是software\\CNKI\\，
        ///       则将判断HKEY_CLASSES_ROOT\\software\\CNKI\\注册表项是否存在
        /// </summary> 
        /// <param name="subKey">注册表项名称（例：software\\CNKI\\）</param>
        /// <param name="regDomain">注册表基项域</param> 
        /// <returns>返回注册表项是否存在，存在返回true，否则返回false</returns>
        /// <remarks>
        /// 2014-02-27 add by yzq
        /// </remarks>
        public bool IsSubKeyExist(string subKey, RegDomain regDomain)
        {
            //判断注册表项名称是否为空，如果为空，返回false 
            if (string.IsNullOrEmpty(subKey))
            {
                return false;
            }
            //检索注册表子项 
            //如果sKey为null,说明没有该注册表项不存在，否则存在 
            RegistryKey sKey = OpenSubKey(subKey, regDomain, false);
            if (sKey == null)
            {
                return false;
            }
            return true;
        }
        /// <summary>
        /// 删除注册表项 
        /// </summary> 
        /// <param name="subKey">注册表项名称（例：software\\CNKI\\）</param>
        /// <param name="regDomain">注册表基项域</param> 
        /// <returns>如果删除成功，则返回true，否则为false</returns>
        /// <remarks>
        /// 2014-02-27 add by yzq
        /// </remarks>
        public bool DeleteSubKey(string subKey, RegDomain regDomain)
        {
            //返回删除是否成功 
            bool result = false;
            //判断注册表项名称是否为空，如果为空，返回false 
            if (string.IsNullOrEmpty(subKey))
            {
                return false;
            }
            //创建基于注册表基项的节点
            RegistryKey key = GetRegDomain(regDomain);
            if (IsSubKeyExist(subKey, regDomain))
            {
                try
                {
                    //删除注册表项 
                    key.DeleteSubKey(subKey);
                    result = true;
                }
                catch
                {
                    result = false;
                }
            }
            //关闭对注册表项的更改 
            key.Close();
            return result;
        }
        #endregion

        #region 键值操作
        /// <summary>
        /// 获取键值信息
        /// </summary>
        /// <param name="subKey">注册项</param>
        /// <param name="regDomain"></param>
        /// <returns></returns>
        /// <remarks>
        /// 2014-02-27 add by yzq
        /// </remarks>
        public Dictionary<string, string> GetRegeditKeyValue(string subKey, RegDomain regDomain)
        {
            Dictionary<string, string> valueDictionary = new Dictionary<string, string>();
            //判断注册表项是否存在
            if (IsSubKeyExist(subKey, regDomain))
            {
                //打开注册表项
                RegistryKey key = OpenSubKey(subKey, regDomain, false);
                //键值集合 
                string[] regeditKeyNames;
                //获取键值集合 
                regeditKeyNames = key.GetValueNames();
                //遍历键值集合，如果存在键值，则退出遍历 
                foreach (string regeditKey in regeditKeyNames)
                {
                    var value = key.GetValue(regeditKey).ToString().Trim();
                    valueDictionary.Add(regeditKey, value);
                }
                //关闭对注册表项的更改
                key.Close();
            }
            return valueDictionary;
        }
        /// <summary>
        /// 判断键值是否存在 
        /// </summary> 
        /// <param name="name">键值名称</param> 
        /// <param name="subKey">注册表项名称（例：software\\CNKI\\）</param> 
        /// <param name="regDomain">注册表基项域</param> 
        /// <returns>返回键值是否存在，存在返回true，否则返回false</returns>
        /// <remarks>
        /// 2014-02-27 add by yzq
        /// </remarks> 
        public bool IsRegeditKeyExist(string name, string subKey, RegDomain regDomain)
        {
            //返回结果 
            bool result = false;
            //判断是否设置键值属性 
            //if (string.IsNullOrEmpty(name))
            //{
            //    return false;
            //}
            //判断注册表项是否存在
            if (IsSubKeyExist(subKey, regDomain))
            {
                //打开注册表项
                RegistryKey key = OpenSubKey(subKey, regDomain, false);
                //键值集合 
                string[] regeditKeyNames;
                //获取键值集合 
                regeditKeyNames = key.GetValueNames();
                //遍历键值集合，如果存在键值，则退出遍历 
                foreach (string regeditKey in regeditKeyNames)
                {
                    if (regeditKey.Equals(name))
                    {
                        result = true;
                        break;
                    }
                }
                //关闭对注册表项的更改
                key.Close();
            }
            return result;
        }
        /// <summary> 
        /// 设置指定的键值内容，指定内容数据类型（请先设置SubKey属性）
        /// 存在改键值则修改键值内容，不存在键值则先创建键值，再设置键值内容 
        /// </summary>
        /// <param name="subKey">注册表项名称（例：software\\CNKI\\）</param> 
        /// <param name="regDomain">注册表基项域</param> 
        /// <param name="name">键值名称</param> 
        /// <param name="content">键值内容</param>
        /// <param name="errMsg">错误信息</param>
        /// <returns>键值内容设置成功，则返回true，否则返回false</returns> 
        /// <remarks>
        /// 2014-02-27 add by yzq
        /// </remarks>
        public bool WriteRegeditKey(string subKey, RegDomain regDomain, string name, object content, out string errMsg)
        {
            //返回结果
            bool result = false;
            errMsg = string.Empty;
            //判断键值是否存在 
            //if (string.IsNullOrEmpty(name))
            //{
            //    errMsg = "未找到指定的键值名称";
            //    return false;
            //}
            //判断注册表项是否存在，如果不存在，则直接创建 
            if (!IsSubKeyExist(subKey, regDomain))
            {
                CreateSubKey(subKey, regDomain);
            }
            //以可写方式打开注册表项 
            RegistryKey key = OpenSubKey(subKey, regDomain, true);
            //如果注册表项打开失败，则返回false 
            if (key == null)
            {
                return false;
            }
            try
            {
                key.SetValue(name, content);
                result = true;
            }
            catch (Exception ex)
            {
                errMsg = "设置键值内容失败，错误信息：" + ex.Message;
                result = false;
            }
            finally
            {
                //关闭对注册表项的更改 
                key.Close();
            }
            return result;
        }
        /// <summary> 
        /// 读取键值内容
        /// </summary>
        /// <param name="name">键值名称</param>
        /// <param name="subKey">注册表项名称</param> 
        /// <param name="regDomain">注册表基项域</param> 
        /// <returns>返回键值内容</returns> 
        /// <remarks>
        /// 2014-02-27 add by yzq
        /// </remarks>
        public object ReadRegeditKey(string name, string subKey, RegDomain regDomain)
        {
            //键值内容结果
            string realValue = null;
            //判断是否设置键值属性 
            //if (string.IsNullOrEmpty(name))
            //{
            //    return null;
            //}
            //判断键值是否存在 
            if (IsRegeditKeyExist(name, subKey, regDomain))
            {
                //打开注册表项 
                RegistryKey key = OpenSubKey(subKey, regDomain, false);
                if (key != null)
                {
                    realValue = key.GetValue(name).ToString().Trim();
                    //关闭对注册表项的更改 
                    key.Close();
                }
            }
            return realValue;
        }
        /// <summary>
        /// 删除键值
        /// </summary>
        /// <param name="name">键值名称</param>
        /// <param name="subKey">注册表项名称</param>
        /// <param name="regDomain">注册表基项域</param>
        /// <param name="errMsg">错误信息</param>
        /// <returns>如果删除成功，返回true，否则返回false</returns>
        /// <remarks>
        /// 2014-02-27 add by yzq
        /// </remarks>
        public bool DeleteRegeditKey(string name, string subKey, RegDomain regDomain, out string errMsg)
        {
            //删除结果
            bool result = false;
            errMsg = string.Empty;
            //判断键值名称和注册表项名称是否为空，如果为空，则返回false
            if (string.IsNullOrEmpty(name) || string.IsNullOrEmpty(subKey))
            {
                errMsg = "未找到指定键值名称";
                return false;
            }
            //判断键值是否存在
            if (IsRegeditKeyExist(name, subKey, regDomain))
            {
                //以可写方式打开注册表项
                RegistryKey key = OpenSubKey(subKey, regDomain, true);
                if (key != null)
                {
                    try
                    {
                        //删除键值
                        key.DeleteValue(name);
                        result = true;
                    }
                    catch (Exception ex)
                    {
                        errMsg = "删除键值失败,错误原因:" + ex.Message;
                        result = false;
                    }
                    finally
                    {
                        //关闭对注册表项的更改
                        key.Close();
                    }
                }
            }
            return result;
        }
        #endregion

        #endregion
        #region 辅助方法
        /// <summary>
        /// 获取注册表基项域对应顶级节点(枚举转换)
        /// 例子：如regDomain是ClassesRoot，则返回Registry.ClassesRoot
        /// </summary>
        /// <param name="regDomain">注册表基项域</param>
        /// <returns>注册表基项域对应顶级节点</returns>
        /// <remarks>
        /// 2014-02-27 add by yzq
        /// </remarks>
        protected RegistryKey GetRegDomain(RegDomain regDomain)
        {
            RegistryKey key;
            switch (regDomain)
            {
                case RegDomain.ClassesRoot:
                    key = Registry.ClassesRoot;
                    break;
                case RegDomain.CurrentUser:
                    key = Registry.CurrentUser;
                    break;
                case RegDomain.LocalMachine:
                    key = Registry.LocalMachine;
                    break;
                case RegDomain.User:
                    key = Registry.Users;
                    break;
                case RegDomain.CurrentConfig:
                    key = Registry.CurrentConfig;
                    break;
                case RegDomain.DynDa:
                    key = Registry.DynData;
                    break;
                case RegDomain.PerformanceData:
                    key = Registry.PerformanceData;
                    break;
                default:
                    key = Registry.LocalMachine;
                    break;
            }
            return key;
        }
        /// <summary>
        /// 打开注册表项节点
        /// </summary>
        /// <param name="subKey">注册表项名称</param>
        /// <param name="regDomain">注册表基项域</param>
        /// <param name="writable">如果需要项的写访问权限，则设置为 true</param>
        /// <returns>如果SubKey为空、null或者SubKey指示注册表项不存在，则返回null，否则返回注册表节点</returns>
        /// <remarks>
        /// 2014-02-27 add by yzq
        /// </remarks>
        protected RegistryKey OpenSubKey(string subKey, RegDomain regDomain, bool writable)
        {
            //判断注册表项名称是否为空
            if (string.IsNullOrEmpty(subKey))
            {
                return null;
            }
            //创建基于注册表基项的节点
            RegistryKey key = GetRegDomain(regDomain);
            //要打开的注册表项的节点
            RegistryKey sKey = null;
            //打开注册表项
            sKey = key.OpenSubKey(subKey, writable);
            //关闭对注册表项的更改
            key.Close();
            //返回注册表节点
            return sKey;
        }
        #endregion

        /// <summary>
        /// 注册表基项静态域
        /// </summary>
        public enum RegDomain
        {
            /// <summary>
            /// 对应于HKEY_CLASSES_ROOT主键
            /// </summary>
            ClassesRoot = 0,
            /// <summary>
            /// 对应于HKEY_CURRENT_USER主键
            /// </summary>
            CurrentUser,
            /// <summary>
            /// 对应于HKEY_LOCAL_MACHINE主键
            /// </summary>
            LocalMachine,
            /// <summary>
            /// 对应于HKEY_USER主键
            /// </summary>
            User,
            /// <summary>
            /// 对应于HEKY_CURRENT_CONFIG主键
            /// </summary>
            CurrentConfig,
            /// <summary>
            /// 对应于HKEY_DYN_DATA主键
            /// </summary>
            DynDa,
            /// <summary>
            /// 对应于HKEY_PERFORMANCE_DATA主键
            /// </summary>
            PerformanceData
        }
    }
}
