using System.Windows;
using System.Windows.Input;

namespace TestInterface.Windows
{
    public interface IMarshalWindow
    {
        string Title { get; }
        void RaiseEvent<T>(string name, RoutedEventArgs args);
        bool Focus<T>(string name);
        void SelectAll<T>(string name);
        void PressKey<T>(string name, Key key);
        void TypeUsingKeyboard<T>(string name, string s);
        bool IsType<T>();

        void Show();
        bool? ShowDialog();
        void Close();
        void LeftClick<T>(string name);
        Point GetListElementPosition<T>(string name, int index);
        Point GetPosition<T>(string name);
    }
}
