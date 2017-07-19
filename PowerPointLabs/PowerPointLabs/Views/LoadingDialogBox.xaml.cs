
namespace PowerPointLabs.Views
{
    /// <summary>
    /// Interaction logic for LoadingDialogBox.xaml
    /// </summary>
    public partial class LoadingDialogBox
    {
        public LoadingDialogBox(string title = TextCollection.LoadingDialogDefaultTitle,
                                string content = TextCollection.LoadingDialogDefaultContent)
        {
            InitializeComponent();
            Title = title;
            contentLabel.Text = content;
        }
    }
}
