using System.Windows;

namespace PowerPointLabs.Views
{
    /// <summary>
    /// Interaction logic for MessageBox.xaml
    /// </summary>
    public partial class MessageBox
    {
        public MessageBox()
        {
            InitializeComponent();
        }

        public MessageBox(string title)
        {
            Title = title;
        }

        public MessageBox(string title, string caption)
        {

        }

        private void OkButton_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }
    }
}
