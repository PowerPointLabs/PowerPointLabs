using PowerPointLabs.TextCollection;

namespace PowerPointLabs.Views
{
    /// <summary>
    /// Interaction logic for LoadingDialogBox.xaml
    /// </summary>
    public partial class LoadingDialogBox
    {
        public LoadingDialogBox(string title = CommonText.LoadingDialogDefaultTitle,
                                string content = CommonText.LoadingDialogDefaultContent)
        {
            InitializeComponent();
            Title = title;
            contentLabel.Text = content;
        }
    }
}
