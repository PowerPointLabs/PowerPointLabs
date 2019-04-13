using System.Windows;
using System.Windows.Controls;

namespace PowerPointLabs.ShapesLab.Views
{
    class CustomMenuItem: MenuItem
    {
        public string actualName;
        public CustomMenuItem(string name, RoutedEventHandler clickHandler)
        {
            actualName = name;
            Header = name;
            Click += clickHandler;
        }
    }
}
