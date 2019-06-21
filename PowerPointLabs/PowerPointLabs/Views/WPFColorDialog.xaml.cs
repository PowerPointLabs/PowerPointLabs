using System.Drawing;
using PowerPointLabs.Utils.Windows;

namespace PowerPointLabs.Views
{
    /// <summary>
    /// Interaction logic for WPFColorDialog.xaml
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
