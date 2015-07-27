using System.Diagnostics;
using System.Windows.Navigation;

namespace PowerPointLabs.ImageSearch
{
    /// <summary>
    /// Interaction logic for SearchOptionsPane.xaml
    /// </summary>
    public partial class SearchOptionsPane
    {
        public SearchOptionsPane()
        {
            InitializeComponent();
        }

        private void Hyperlink_OnRequestNavigate(object sender, RequestNavigateEventArgs e)
        {
            Process.Start(new ProcessStartInfo(e.Uri.AbsoluteUri));
            e.Handled = true;
        }
    }
}
