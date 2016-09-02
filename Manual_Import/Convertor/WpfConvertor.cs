﻿using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Windows.Data;
using Manual_Import.Model;
using System.Drawing;

namespace Manual_Import.Convertor
{
    [ValueConversion(typeof(int), typeof(string))]
    class HasTidyConvertor : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            int hasTidy = System.Convert.ToInt32(value);
            if (hasTidy == 1)
                return "Resources/ok.ico";
            else if (hasTidy == 2)
                return "Resources/upload.ico";
            else if (hasTidy == -1)
                return "Resources/error.ico";
            else if (hasTidy == 0)
                return "Resources/not.ico";
            else
                return null;
        }
        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            return true;
        }
    }
    [ValueConversion(typeof(SystemType), typeof(string))]
    class SystemTypeConvertor : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            SystemType st = (SystemType)value;
            if (st == SystemType.Dir)
                return "Resources/dir.ico";
            else if (st == SystemType.PDF)
                return "Resources/pdf.ico";
            else if (st == SystemType.Word)
                return "Resources/word.ico";
            else
                return "Resources/file.ico";
        }
        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            return true;
        }
        
    }
    [ValueConversion(typeof(bool), typeof(System.Windows.Media.Brush))]
    class ForegroundConvertor : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            bool ischecked = (bool)value;
            if (ischecked)
                return System.Windows.Media.Brushes.Black;
            else
                return System.Windows.Media.Brushes.Gray;
        }
        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            return true;
        }
    }
}
