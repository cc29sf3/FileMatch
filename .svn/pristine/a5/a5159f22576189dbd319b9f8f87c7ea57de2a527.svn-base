using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.ComponentModel; 

namespace Manual_Import.Model
{
    public class Model_FileSystem : INotifyPropertyChanged 
    {
        /// <summary>
        /// 是否勾选
        /// </summary>
        public bool Checked
        {
            get { return _checked; }
            set {
                if (!value)
                {
                    _checked = value;
                    RaisePropertyChanged("Checked");
                }
                else if (value && (HasTidy == 0 || HasTidy == -2))
                {
                    _checked = value;
                    RaisePropertyChanged("Checked");               
                }
               
            }
        }
        /// <summary>
        /// 文件或文件夹名称
        /// </summary>
        public string Name
        {
            get { return _name; }
            set {
                _name = value;
                //RaisePropertyChanged("Name");
            }
        }
        /// <summary>
        /// 文件路径
        /// </summary>
        public string FullPath
        {
            get { return _fullpath; }
            set {
                _fullpath = value;
                //RaisePropertyChanged("FullPath");
            }
        }
        /// <summary>
        /// 文件夹或文件
        /// </summary>
        public SystemType Type
        {
            get { return _type; }
            set
            {
                _type = value;
                //RaisePropertyChanged("Type");
            }
        }
        /// <summary>
        /// 该文件目前状态
        /// 未整理:0 整理成功:1 整理失败:-1 篇提交:2 文件夹:-2
        /// </summary>
        public int HasTidy
        {
            get { return _hasTidy; }
            set
            {
                _hasTidy = value;
                RaisePropertyChanged("HasTidy");
            }
        }
        /// <summary>
        /// 文件大小
        /// </summary>
        public string FileSize
        {
            get { return _filesize; }
            set { _filesize = value; }
        }
        /// <summary>
        /// 文件后缀名
        /// </summary>
        public string Extension
        {
            get { return _extension; }
            set { _extension = value; }
        }
        /// <summary>
        /// 文件修改时间
        /// </summary>
        public string Time
        {
            get { return _time; }
            set { _time = value; }
        }

        private bool _checked;
        private string _name;
        private string _fullpath;
        private SystemType _type;
        private int _hasTidy;
        private string _filesize;
        private string _extension;
        private string _time;

        public event PropertyChangedEventHandler PropertyChanged;
        public void RaisePropertyChanged(string Propertyname)
        {
            if (PropertyChanged != null)
            {
                this.PropertyChanged(this, new PropertyChangedEventArgs(Propertyname));
            }
        }
    }
}
