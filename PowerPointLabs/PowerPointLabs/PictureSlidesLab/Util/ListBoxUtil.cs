using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;

namespace PowerPointLabs.PictureSlidesLab.Util
{
    public class ListBoxUtil
    {
        public static ScrollViewer FindScrollViewer(DependencyObject parent)
        {
            int childCount = VisualTreeHelper.GetChildrenCount(parent);
            for (int i = 0; i < childCount; i++)
            {
                DependencyObject elt = VisualTreeHelper.GetChild(parent, i);
                if (elt is ScrollViewer)
                {
                    return (ScrollViewer)elt;
                }

                ScrollViewer result = FindScrollViewer(elt);
                if (result != null)
                {
                    return result;
                }
            }
            return null;
        }
    }
}
