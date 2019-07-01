using System.ComponentModel;
using System.Windows;
using System.Windows.Input;

using MahApps.Metro.Controls;

using PowerPointLabs.PictureSlidesLab.Util;
using PowerPointLabs.Utils;

namespace PowerPointLabs.PictureSlidesLab.Views
{
    /// <summary>
    /// Interaction logic for QuickDropDialog.xaml
    /// </summary>
    public partial class QuickDropDialog
    {
        public delegate void DropHandle(object sender, DragEventArgs e);

        public event DropHandle DropHandler;

        public delegate void PasteHandle();

        public event PasteHandle PasteHandler;

        // indicate whether the window is open/closed or not
        public bool IsOpen { get; set; }

        private MetroWindow _parent;

        public bool IsDisposed
        {
            get
            {
                System.Reflection.PropertyInfo propertyInfo = typeof(Window).GetProperty("IsDisposed",
                    System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance);
                return (bool) propertyInfo.GetValue(this, null);
            }
        }

        public QuickDropDialog(MetroWindow parent)
        {
            _parent = parent;
            InitializeComponent();

            InitDragAndDrop();
            IsOpen = true;
            PictureSlidesLabLogo.Source = GraphicsUtil.BitmapToImageSource(Properties.Resources.PictureSlidesLab);
        }

        public void ShowQuickDropDialog()
        {
            IsOpen = true;
            Show();
        }

        public void HideQuickDropDialog()
        {
            IsOpen = false;
            Hide();
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
            if (_parent.WindowState == WindowState.Minimized)
            {
                _parent.WindowState = WindowState.Normal;
            }
            _parent.Activate();
            e.Handled = true;
        }

        private void MenuItemPastePictureHere_OnClick(object sender, RoutedEventArgs e)
        {
            if (PasteHandler != null)
            {
                PasteHandler();
            }
        }

        private void QuickDropDialog_OnKeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.V
                && Keyboard.Modifiers == ModifierKeys.Control
                && PasteHandler != null)
            {
                PasteHandler();
            }
        }
    }
}
