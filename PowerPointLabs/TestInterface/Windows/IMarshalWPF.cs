using System;
using System.Windows;

namespace TestInterface.Windows
{
    public interface IMarshalWPF
    {
        string Title { get; }
        Type Type { get; }
        void RaiseEvent(string name, RoutedEventArgs args);
        bool Focus(string name);
        void Close();
    }
}
