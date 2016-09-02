using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using Utility.Log;

namespace Utility.Common
{
    /// <summary>
    /// 描述：文件相关操作类
    /// 作者：员战强
    /// 日期：2014-02-28
    /// </summary>
    public static class FileManage
    {
        /// <summary>
        /// 检测文件是否可以
        /// </summary>
        /// <param name="path"></param>
        /// <returns></returns>
        public static bool IsCanWrite(this string path)
        {
            bool isCanWrite = false;
            FileStream fs = null;
            try
            {
                fs = new FileStream(path, FileMode.OpenOrCreate, FileAccess.Write);
                isCanWrite = fs.CanWrite;
            }
            catch (Exception)
            {
                isCanWrite = false;
            }
            finally
            {
                if (fs!=null)
                {
                    fs.Close();
                }
            }
            return isCanWrite;
        }
        /// <summary>
        /// 是否含有文件夹
        /// </summary>
        /// <param name="path">文件夹路径</param>
        /// <returns></returns>
        /// <remarks>
        /// 2014-02-28 add by yzq
        /// </remarks>
        public static bool IsHasDirectory(this string path)
        {
            return Directory.Exists(path);
        }
        /// <summary>
        /// 是否含有文件
        /// </summary>
        /// <param name="file">文件夹</param> 
        /// <returns></returns>
        /// <remarks>
        /// 2014-02-28 add by yzq
        /// </remarks>
        public static bool IsHasFile(this string file)
        {
            return File.Exists(file);
        }
        /// <summary>
        /// 移动文件夹到备份路径
        /// </summary>
        /// <param name="movePath">需要移动的文件</param>
        /// <param name="newPath">移动的位置</param>
        /// <param name="errMsg">返回的错误信息</param>
        /// <returns></returns>
        /// <remarks>
        /// 2014-03-05 add by yzq
        /// </remarks>
        public static bool MoveDirectory(this string movePath, string newPath, out string errMsg)
        {
            errMsg = "";
            bool isMoveAll = true;
            try
            {
                //Directory.Move(movePath, newPath);
                //移动文件夹下文件
                string[] files = Directory.GetFiles(movePath);
                foreach (string file in files)
                {
                    var toFilePath = Path.Combine(newPath, Path.GetFileName(file));
                    try
                    {
                        if (!file.MoveFile(toFilePath, out errMsg))
                        {
                            isMoveAll = false;
                            TextLog.WritwLog("移动文件" + file + "出错，错误信息：" + errMsg);
                        }
                    }
                    catch (Exception ex)
                    {
                        isMoveAll = false;
                        TextLog.WritwLog("移动文件" + file + "出错，错误信息：" + ex.Message);
                    }
                }
                string[] dirs = Directory.GetDirectories(movePath);
                foreach (string dir in dirs)
                {
                    var childPath = dir.ToLower().Replace(movePath.ToLower(), newPath.ToLower());
                    if (!Directory.Exists(childPath))
                    {
                        Directory.CreateDirectory(childPath);
                    }
                    var moveResult = MoveDirectory(dir, childPath, out errMsg);
                    if (isMoveAll)
                    {
                        isMoveAll = moveResult;
                    }
                }
                Directory.Delete(movePath, true);
            }
            catch (Exception ex)
            {
                isMoveAll = false;
                errMsg = ex.Message;
            }
            return isMoveAll;
        }
        /// <summary>
        /// 移动文件
        /// </summary>
        /// <param name="movePath"></param>
        /// <param name="newPath"></param>
        /// <param name="errMsg"></param>
        /// <returns></returns>
        public static bool MoveFile(this string movePath, string newPath, out string errMsg)
        {
            errMsg = "";
            try
            {
                File.Copy(movePath, newPath);
                File.Delete(movePath);
                return true;
            }
            catch (Exception ex)
            {
                errMsg = ex.Message;
                return false;
            }
        }
        /// <summary>
        /// 创建文件夹
        /// </summary>
        /// <param name="createPath">文件夹路径</param>
        /// <param name="errMsg">返回的消息</param>
        /// <returns></returns>
        /// <remarks>
        /// 2014-03-05 add by yzq
        /// </remarks>
        public static bool CreateDirectory(this string createPath, out string errMsg)
        {
            errMsg = "";
            if (!createPath.IsHasDirectory())
            {
                try
                {
                    Directory.CreateDirectory(createPath);
                    return true;
                }
                catch (Exception ex)
                {
                    errMsg = ex.Message;
                    return false;
                }
            }
            return true;
        }
        /// <summary>
        /// 删除指定的工作文件路径
        /// </summary>
        /// <param name="delPath">需要删除的文件</param>
        /// <param name="errMsg">返回的消息</param>
        /// <returns></returns>
        /// <remarks>
        /// 2014-03-05 add by yzq
        /// </remarks>
        public static bool DeleteDirectory(this string delPath, out string errMsg)
        {
            errMsg = "";
            try
            {
                if (delPath.IsHasDirectory())
                {
                    Directory.Delete(delPath, true);
                }
                return true;
            }
            catch (Exception ex)
            {
                errMsg = ex.Message;
                return false;
            }
        }
        /// <summary>
        /// 删除文件
        /// </summary>
        /// <param name="delPath">文件路径</param>
        /// <param name="errMsg">错误信息</param>
        /// <returns></returns>
        public static bool FileDirectory(this string delPath, out string errMsg)
        {
            errMsg = "";
            try
            {
                if (delPath.IsHasFile())
                {
                    File.Delete(delPath);
                }
                return true;
            }
            catch (Exception ex)
            {
                errMsg = ex.Message;
                return false;
            }
        }
        /// <summary>
        /// 获取需要上传的文件信息
        /// </summary>
        /// <param name="path">路径</param>
        /// <param name="uploadList">上传文件集合</param>
        /// <returns></returns>
        public static void GetUploadList(string path, ref IList<string> uploadList)
        {
            if (path.IsHasDirectory())
            {
                string[] files = Directory.GetFiles(path);
                foreach (string file in files)
                {
                    uploadList.Add(file);
                }
                string[] dirs = Directory.GetDirectories(path);
                foreach (var dir in dirs)
                {
                    GetUploadList(dir, ref uploadList);
                }
            }
        }
        /// <summary>
        /// 本地工作路径（相对路径，不包含根路径）
        /// </summary>
        /// <param name="lineID">工作线编号</param>
        /// <param name="postID">岗位编号</param>
        /// <param name="taskCode">任务编号</param>
        /// <param name="definePath">自定义文件夹</param>
        /// <returns></returns>
        public static string WorkPath(int lineID, int postID, string taskCode, string definePath)
        {
            return definePath + @"\" + DateTime.Now.ToString("yyyyMMdd") + @"\" + DateTime.Now.ToFileTimeUtc() + @"\" + lineID + @"\" + postID + @"\" + taskCode;
        }

        public static void copyDirectory(string sourceDirectory, string destDirectory)
        {
            //判断源目录和目标目录是否存在，如果不存在，则创建一个目录
            if (!Directory.Exists(sourceDirectory))
            {
                Directory.CreateDirectory(sourceDirectory);
            }
            if (!Directory.Exists(destDirectory))
            {
                Directory.CreateDirectory(destDirectory);
            }
            //拷贝文件
            copyFile(sourceDirectory, destDirectory);

            //拷贝子目录       
            //获取所有子目录名称
            string[] directionName = Directory.GetDirectories(sourceDirectory);

            foreach (string directionPath in directionName)
            {
                //根据每个子目录名称生成对应的目标子目录名称
                string directionPathTemp = destDirectory + "\\" + directionPath.Substring(sourceDirectory.Length + 1);

                //递归下去
                copyDirectory(directionPath, directionPathTemp);
            }
        }

        public static void copyFile(string sourceDirectory, string destDirectory)
        {
            //获取所有文件名称
            string[] fileName = Directory.GetFiles(sourceDirectory);

            foreach (string filePath in fileName)
            {
                //根据每个文件名称生成对应的目标文件名称
                string filePathTemp = destDirectory + "\\" + filePath.Substring(sourceDirectory.Length + 1);

                //若不存在，直接复制文件；若存在，覆盖复制
                if (File.Exists(filePathTemp))
                {
                    File.Copy(filePath, filePathTemp, true);
                }
                else
                {
                    File.Copy(filePath, filePathTemp);
                }
            }
        }    
    }
}
