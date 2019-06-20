using System.Drawing;
using PowerPointLabs.Utils;

namespace PowerPointLabs.Views
{
    /// <summary>
    /// Interaction logic for UserControl1.xaml
    /// </summary>
    public partial class WPFColorDialog
    {
        public bool FullOpen { get; set; }
        public Color Color { get; set; }

        public WPFColorDialog()
        {
            InitializeComponent();
        }

        public DialogResult ShowDialog()
        {
            return DialogResult.None;
        }
    }
}
