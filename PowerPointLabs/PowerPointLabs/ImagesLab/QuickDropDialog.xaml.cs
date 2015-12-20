using System.ComponentModel;
using System.Windows;
using System.Windows.Input;
using MahApps.Metro.Controls;
using PowerPointLabs.ImagesLab.Util;

namespace PowerPointLabs.ImagesLab
{
    /// <summary>
    /// Interaction logic for QuickDropDialog.xaml
    /// </summary>
    public partial class QuickDropDialog
    {
        public delegate void DropHandle(object sender, DragEventArgs e);

        public event DropHandle DropHandler;

        // indicate whether the window is open/closed or not
        public bool IsOpen { get; set; }

        private MetroWindow _parent;

        public QuickDropDialog(MetroWindow parent)
        {
            _parent = parent;
            InitializeComponent();

            InitDragAndDrop();
            IsOpen = true;
            ImagesLabLogo.Source = ImageUtil.BitmapToImageSource(Properties.Resources.ImagesLab);
        }

        // drag to move window
        private void QuickDropDialog_OnMouseDown(object sender, MouseButtonEventArgs e)
        {
            if (e.LeftButton == MouseButtonState.Pressed)
            {
                DragMove();
            }
        }

        private void InitDragAndDrop()
        {
            AllowDrop = true;

            DragEnter += OnDragEnter;
            DragLeave += OnDragLeave;
            DragOver += OnDragEnter;
            Drop += OnDrop;
        }

        private void OnDrop(object sender, DragEventArgs e)
        {
            try
            {
                if (DropHandler != null)
                {
                    DropHandler(sender, e);
                }
            }
            finally
            {
                Overlay.Visibility = Visibility.Hidden;
            }
        }

        private void OnDragLeave(object sender, DragEventArgs args)
        {
            Overlay.Visibility = Visibility.Hidden;
        }

        private void OnDragEnter(object sender, DragEventArgs args)
        {
            if (args.Data.GetDataPresent("FileDrop")
                || args.Data.GetDataPresent("Text"))
            {
                Overlay.Visibility = Visibility.Visible;
                Activate();
            }
        }

        private void QuickDropDialog_OnClosing(object sender, CancelEventArgs e)
        {
            IsOpen = false;
        }

        private void QuickDropDialog_OnMouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            _parent.Activate();
            Close();
        }
    }
}
