
namespace PowerPointLabs.Views
{
    /// <summary>
    /// Interaction logic for LoadingDialogBox.xaml
    /// </summary>
    public partial class LoadingDialogBox
    {
        public LoadingDialogBox(string title = TextCollection1.LoadingDialogDefaultTitle,
                                string content = TextCollection1.LoadingDialogDefaultContent)
        {
            InitializeComponent();
            Title = title;
            contentLabel.Text = content;
        }
    }
}
