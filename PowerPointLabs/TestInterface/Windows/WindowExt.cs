using System.Windows;
using System.Windows.Threading;

namespace TestInterface.Windows
{
    public class WindowExt : IWindow
    {
        private Window window;
        public WindowExt(Window w)
        {
            window = w;
        }

        public Dispatcher Dispatcher => window.Dispatcher;

        public string Title => window.Title;

        public void Close()
        {
            window.Close();
        }

        public void Show()
        {
            window.Show();
        }

        public bool? ShowDialog()
        {
            return window.ShowDialog();
        }
    }
}
