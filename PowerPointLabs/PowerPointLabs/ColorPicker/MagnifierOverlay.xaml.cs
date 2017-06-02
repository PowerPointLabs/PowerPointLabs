using System.Windows;
using System.Windows.Interop;

using PPExtraEventHelper;

namespace PowerPointLabs.ColorPicker
{
    /// <summary>
    /// Interaction logic for MagnifierOverlay.xaml
    /// </summary>
    public partial class MagnifierOverlay : Window
    {
        public MagnifierOverlay()
        {
            InitializeComponent();
        }

        private void MagnifierOverlay_Loaded(object sender, RoutedEventArgs e)
        {
            WindowInteropHelper wndHelper = new WindowInteropHelper(this);
            int extendedStyle = Native.GetWindowLong(wndHelper.Handle, (int)Native.WindowLong.GWL_EXSTYLE);

            Native.SetWindowLong(wndHelper.Handle, (int)Native.WindowLong.GWL_EXSTYLE,
                                extendedStyle | (int)Native.ExtendedWindowStyles.WS_EX_TOOLWINDOW);
        }

        private void MagnifierOverlay_Deactivated(object sender, System.EventArgs e)
        {
            PPMouse.StopAllActions();
        }
    }


}
