using Manual_Import.Model;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Text;

namespace Manual_Import.ViewModel
{
    public class ViewModel_Main:INotifyPropertyChanged
    {
        /// <summary>
        /// listview数据项
        /// </summary>
        public ObservableCollection<Model_FileSystem> Models
        {
            get
            {
                return _models;
            }
            set
            {
                _models = value;
                RaisePropertyChanged("Models");
            }
        }
        /// <summary>
        /// 当前用户工号
        /// </summary>
        public string GongHao { get; set; }
        /// <summary>
        /// 用户点击进入的当前路径
        /// </summary>
        public string CurPath
        {
            get { return _curPath; }
            set { _curPath = value; RaisePropertyChanged("CurPath"); }
        }
        /// <summary>
        /// 任务类型,纯盘或刊盘
        /// </summary>
        public string TaskType
        {
            get { return _taskType; }
            set
            {
                _taskType = value;
                RaisePropertyChanged("TaskType");
            }
        }
        /// <summary>
        /// 整理后的xml文件保存路径
        /// </summary>
        public string Xml_Temp_Path{get;set;}
        /// <summary>
        /// 提取的非正文页临时存储路径
        /// </summary>
        public string tempPath { get; set; }
        /// <summary>
        /// 上传路径
        /// </summary>
        public string Upload_Path { get; set; }
        /// <summary>
        /// 起始流水号/编号
        /// </summary>
        public string Begin_Code { get; set; }
        /// <summary>
        /// 待整理文件夹根路径
        /// </summary>
        public string RootPath { get; set; }
        /// <summary>
        /// 整理后文件夹路径
        /// </summary>
        public string AfterTydyPath { get; set; }
        /// <summary>
        /// 任务编号
        /// </summary>
        public string TaskCode
        {
            get { return _taskCode; }
            set
            {
                _taskCode = value;
                RaisePropertyChanged("TaskCode");
            }
        }

        /// <summary>
        /// 工作路径
        /// </summary>
        public string WorkPath { get; set; }
        /// <summary>
        /// 累计整理成功数
        /// </summary>
        public int TotalSuc
        {
            get { return _totalSuc; }
            set { _totalSuc = value; RaisePropertyChanged("TotalSuc"); }
        }
        /// <summary>
        /// 累计整理失败数
        /// </summary>
        public int TotalFail
        {
            get { return _totalFail; }
            set { _totalFail = value; RaisePropertyChanged("TotalFail"); }
        }
        /// <summary>
        /// 本次整理成功数
        /// </summary>
        public int CurSuc
        {
            get { return _curSuc; }
            set { _curSuc = value; RaisePropertyChanged("CurSuc"); }
        }
        /// <summary>
        /// 本次整理失败数
        /// </summary>
        public int CurFail
        {
            get { return _curFail; }
            set { _curFail = value; RaisePropertyChanged("CurFail"); }
        }
        /// <summary>
        /// 全部文件总数
        /// </summary>
        public int Total
        {
            get{return _total;}
            set { 
                _total = value;
                RaisePropertyChanged("Total");
            }
        }
        /// <summary>
        /// 未整理文件数量
        /// </summary>
        public int UnTidy
        {
            get { return _unTidy; }
            set
            {
                _unTidy = value;
                RaisePropertyChanged("UnTidy");
            }
        }
        /// <summary>
        /// 是否禁用checkbox
        /// </summary>
        public bool CheckBoxEnable
        {
            get { return _checkBoxEnable; }
            set
            {
                _checkBoxEnable = value;
                RaisePropertyChanged("CheckBoxEnable");
            }
        }

        private ObservableCollection<Model_FileSystem> _models = new ObservableCollection<Model_FileSystem>();
        private string _curPath;
        private string _taskType;
        private string _taskCode;
        private int _totalSuc;
        private int _totalFail;
        private int _curSuc;
        private int _curFail;
        private int _unTidy;
        private int _total;
        private bool _checkBoxEnable=true;

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
