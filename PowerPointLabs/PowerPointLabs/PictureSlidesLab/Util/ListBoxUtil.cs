using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;

namespace PowerPointLabs.PictureSlidesLab.Util
{
    public class ListBoxUtil
    {
        public static ScrollViewer FindScrollViewer(DependencyObject parent)
        {
            var childCount = VisualTreeHelper.GetChildrenCount(parent);
            for (var i = 0; i < childCount; i++)
            {
                var elt = VisualTreeHelper.GetChild(parent, i);
                if (elt is ScrollViewer)
                {
                    return (ScrollViewer)elt;
                }

                var result = FindScrollViewer(elt);
                if (result != null)
                {
                    return result;
                }
            }
            return null;
        }
    }
}
