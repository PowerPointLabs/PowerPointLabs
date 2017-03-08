using System.Windows;

namespace PowerPointLabs.ColorPicker
{
    /// <summary>
    /// Interaction logic for MagnifierOverlay.xaml
    /// </summary>
    public partial class MagnifierOverlay : Window
    {
        private Point halfsize;

        public MagnifierOverlay()
        {
            InitializeComponent();
            halfsize = new Point(Width / 2, Height / 2);
        }

        public Point HalfSize
        {
            get { return halfsize; }
        }
    }
}
