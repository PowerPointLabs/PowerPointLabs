using System;
using System.Windows.Forms;

namespace PowerPointLabs.Views
{
    public partial class LoadingDialog : Form
    {
        public LoadingDialog(string title = TextCollection.LoadingDialogDefaultTitle,
                             string content = TextCollection.LoadingDialogDefaultContent)
        {
            InitializeComponent();

            Text = title;
            contentLabel.Text = content;

            Width = Math.Max(contentLabel.Width + 2*contentLabel.Location.X, ClientSize.Width);
        }
    }
}
