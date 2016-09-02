using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;

namespace Manual_Import.Helper
{
    public static class ExtendMethod
    {
        public static List<T> GetItemAt<T>(this ListView listview, Rect areaOfInterest)
        {
            var list = new List<T>();
            var rect = new RectangleGeometry(areaOfInterest);
            var hitTestParams = new GeometryHitTestParameters(rect);
            var resultCallback = new HitTestResultCallback(x => HitTestResultBehavior.Continue);
            var filterCallback = new HitTestFilterCallback(x =>
            {
                if (x is ListViewItem)
                {

                    var item = (T)((ListViewItem)x).Content;
                    list.Add(item);
                }
                return HitTestFilterBehavior.Continue;
            });

            VisualTreeHelper.HitTest(listview, filterCallback, resultCallback, hitTestParams);
            return list;
        }
    }
}
