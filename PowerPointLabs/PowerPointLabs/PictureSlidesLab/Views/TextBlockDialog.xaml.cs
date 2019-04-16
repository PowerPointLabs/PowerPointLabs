using System.Windows;

using PowerPointLabs.WPF.Observable;

namespace PowerPointLabs.PictureSlidesLab.Views
{
    /// <summary>
    /// Interaction logic for SlideSelectionDialog.xaml
    /// </summary>
    public partial class TextBlockDialog
    {
        public delegate void OkEvent();

        public event OkEvent OnOkButtonClick;

        public ObservableString DialogTitleProperty { get; set; }

        public ObservableString DialogTextBlockProperty { get; set; }

        public bool IsOpen { get; set; }

        public TextBlockDialog()
        {
            InitializeComponent();
            DialogTitleProperty = new ObservableString();
            DialogTextBlockProperty = new ObservableString();
            DialogTitle.DataContext = DialogTitleProperty;
            DialogTextBlock.DataContext = DialogTextBlockProperty;
        }

        public void OpenDialog()
        {
            IsOpen = true;
        }

        public void CloseDialog()
        {
            IsOpen = false;
        }

        private void OkButton_OnClick(object sender, RoutedEventArgs e)
        {
            if (OnOkButtonClick != null)
            {
                OnOkButtonClick();
            }
        }
    }
}
