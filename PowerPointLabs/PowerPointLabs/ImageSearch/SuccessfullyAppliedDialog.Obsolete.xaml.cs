using System.Windows;
using System.Windows.Input;

namespace PowerPointLabs.ImageSearch
{
    /// <summary>
    /// Interaction logic for SuccessfullyAppliedDialog.xaml
    /// </summary>
    public partial class SuccessfullyAppliedDialog
    {
        public delegate void GotoNextSlideEvent();

        public delegate void OkEvent();

        public event GotoNextSlideEvent OnGotoNextSlide;

        public event OkEvent OnOk;

        public SuccessfullyAppliedDialog()
        {
            InitializeComponent();
        }

        private void GotoNextSlideButton_OnClick(object sender, RoutedEventArgs e)
        {
            if (OnGotoNextSlide != null)
            {
                OnGotoNextSlide();
            }
        }

        private void OkButton_OnClick(object sender, RoutedEventArgs e)
        {
            if (OnOk != null)
            {
                OnOk();
            }
        }

        public void ShowGotoNextSlideButton()
        {
            GotoNextSlideButton.Visibility = Visibility.Visible;
        }

        public void HideGotoNextSlideButton()
        {
            GotoNextSlideButton.Visibility = Visibility.Hidden;
        }

        public void FocusOkButton()
        {
            OkButton.Focusable = true;
            OkButton.Focus();
        }

        public void FocusGotoNextSlideButton()
        {
            GotoNextSlideButton.Focusable = true;
            GotoNextSlideButton.Focus();
        }
    }
}
