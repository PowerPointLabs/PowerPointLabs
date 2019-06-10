using System.Windows.Threading;

namespace TestInterface.Windows
{
    public interface IWindow
    {
        Dispatcher Dispatcher { get; }
        string Title { get; }
        void Show();
        bool? ShowDialog();
        void Close();
        // TODO: Add more actions to support more generic operations
    }
}
