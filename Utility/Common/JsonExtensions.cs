using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.Serialization.Formatters.Binary;
using System.Text;
using System.Xml;
using System.Xml.Serialization;
using System.Runtime.Serialization.Json;

namespace Utility.Common
{
    /// <summary>
    /// 描述：json序列化对象扩展方法
    /// 作者：员战强
    /// 日期：2014-02-27
    /// </summary>
    public static class JsonExtensions
    {
        /// <summary>
        /// 将对象转换为json字符串
        /// </summary>
        /// <param name="obj">对象</param>
        /// <returns></returns>
        /// <remarks>
        /// 2014-02-27 add by yzq
        /// </remarks>
        public static string ToJson(this object obj)
        {
            using (var ms = new MemoryStream())
            {
                new DataContractJsonSerializer(obj.GetType()).WriteObject(ms, obj);
                return Encoding.UTF8.GetString(ms.ToArray());
            }
        }
        /// <summary>
        /// 将json字符串转换为对象
        /// </summary>
        /// <typeparam name="TResult">对象类型</typeparam>
        /// <param name="json">json字符串</param>
        /// <returns></returns>
        /// <remarks>
        /// 2014-02-27 add by yzq
        /// </remarks>
        public static TResult FromJson<TResult>(this string json)
        {
            using (var ms = new MemoryStream(Encoding.UTF8.GetBytes(json)))
            {
                return (TResult)new DataContractJsonSerializer(typeof(TResult)).ReadObject(ms);
            }
        }
        /// <summary>
        /// 将对象转换为xml文本字符
        /// </summary>
        /// <param name="obj">对象</param>
        /// <returns></returns>
        /// <remarks>
        /// 2014-02-27 add by yzq
        /// </remarks>
        public static string ToXml(this object obj)
        {
            var ser = new XmlSerializer(obj.GetType());
            var stream = new MemoryStream();
            ser.Serialize(stream, obj);
            return Encoding.UTF8.GetString(stream.ToArray());
        }
        /// <summary>
        /// 将xml文本字符转换为对象
        /// </summary>
        /// <typeparam name="TResult">对象类型</typeparam>
        /// <param name="xml">xml内容</param>
        /// <returns></returns>
        /// <remarks>
        /// 2014-02-27 add by yzq
        /// </remarks>
        public static TResult FromXml<TResult>(this string xml)
        {
            var ser = new XmlSerializer(typeof(TResult));
            var ms = new MemoryStream(Encoding.Default.GetBytes(xml));

            return (TResult)ser.Deserialize(ms);
        }
        /// <summary>
        /// 将对象转换成xml，并储存在硬盘
        /// </summary>
        /// <param name="obj">对象内容</param>
        /// <param name="path">储存的文件地址</param>
        /// <remarks>
        /// 2014-02-27 add by yzq
        /// </remarks>
        public static void SaveToXml(this object obj, string path)
        {
            try
            {
                XmlSerializer serializer = new XmlSerializer(obj.GetType());
                using (
                    Stream stream = new FileStream(path, FileMode.Create, FileAccess.Write, FileShare.None)
                    )
                {
                    XmlSerializerNamespaces namespaces = new XmlSerializerNamespaces();
                    namespaces.Add("", "");
                    serializer.Serialize(stream, obj, namespaces);
                }
            }
            catch (Exception ex)
            {
                var a = ex.Message;
            }
        }
        /// <summary>
        /// 从硬盘读取xml文件并转换成对象
        /// </summary>
        /// <typeparam name="TResult">对象类型</typeparam>
        /// <param name="path">文本路径</param>
        /// <returns></returns>
        /// <remarks>
        /// 2014-02-27 add by yzq
        /// </remarks>
        public static TResult LoadFromXml<TResult>(this string path)
        {
            var ser = new XmlSerializer(typeof(TResult));
            var xmlreader = new XmlTextReader(path);
            try
            {
                if (!ser.CanDeserialize(xmlreader))
                {
                    throw new ArgumentException("无法反序列化为指定类型," + path + " ->" + typeof(TResult));
                }
                return (TResult)ser.Deserialize(xmlreader);
            }
            finally
            {
                xmlreader.Close();
            }
        }
        /// <summary>
        /// 序列化信息到文本中
        /// </summary>
        /// <param name="obj">对象</param>
        /// <param name="path">路径</param>
        /// <remarks>
        /// 2014-02-27 add by yzq
        /// </remarks>
        public static void SaveToBin(this object obj, string path)
        {
            using (FileStream stream = new FileStream(path, FileMode.Create))
            {
                BinaryFormatter ser = new BinaryFormatter();
                ser.Serialize(stream, obj);
            }
        }
        /// <summary>
        /// 从硬盘读取bin文件并转换成对象
        /// </summary>
        /// <typeparam name="TResult">对象类型</typeparam>
        /// <param name="path">文本路径</param>
        /// <returns></returns>
        /// <remarks>
        /// 2014-02-27 add by yzq
        /// </remarks>
        public static TResult LoadFromBin<TResult>(this string path)
        {
            using (FileStream stream = new FileStream(path, FileMode.OpenOrCreate))
            {
                var ser = new BinaryFormatter();
                try
                {
                    return (TResult)ser.Deserialize(stream);
                }
                catch (Exception)
                {
                    throw new ArgumentException("无法反序列化为指定类型," + path + " ->" + typeof(TResult));
                }
            }
        }
    }
}
