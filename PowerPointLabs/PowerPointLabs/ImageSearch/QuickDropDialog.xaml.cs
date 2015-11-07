using System.ComponentModel;
using System.Windows;
using System.Windows.Input;
using MahApps.Metro.Controls;
using PowerPointLabs.ImageSearch.Domain;
using PowerPointLabs.ImageSearch.Util;

namespace PowerPointLabs.ImageSearch
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

        private const string DialogSettingsFilename = "quick-drop-dialog.xml";

        public QuickDropDialog(MetroWindow parent)
        {
            _parent = parent;
            InitializeComponent();

            // init window pos
            var windowInfo = StoragePath.LoadWindowInfo(DialogSettingsFilename);
            if (windowInfo.Left != -1)
            {
                Left = windowInfo.Left;
                Top = windowInfo.Top;
            }
            else
            {
                Left = SystemParameters.PrimaryScreenWidth - Width - 100;
                Top = SystemParameters.PrimaryScreenHeight - Height - 50;
            }

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

            var windowInfo = new WindowInfo();
            windowInfo.Left = Left;
            windowInfo.Top = Top;
            StoragePath.Save(DialogSettingsFilename, windowInfo);
        }

        private void QuickDropDialog_OnMouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            _parent.Activate();
            Close();
        }
    }
}
