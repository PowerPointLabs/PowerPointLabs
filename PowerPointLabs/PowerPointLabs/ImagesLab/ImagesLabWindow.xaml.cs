using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Collections.Specialized;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Input;
using System.Windows.Media;
using PowerPointLabs.ImagesLab.Domain;
using PowerPointLabs.ImagesLab.Handler;
using PowerPointLabs.ImagesLab.Util;
using PowerPointLabs.Models;
using ButtonBase = System.Windows.Controls.Primitives.ButtonBase;
using KeyEventArgs = System.Windows.Input.KeyEventArgs;
using ListBox = System.Windows.Controls.ListBox;

namespace PowerPointLabs.ImagesLab
{
    /// <summary>
    /// Interaction logic for Images Lab
    /// </summary>
    public partial class ImagesLabWindow
    {
        # region Props & States

        // UI model - list that holds image items
        public ObservableCollection<ImageItem> ImageSelectionList { get; set; }

        // UI model - list that holds styles preview items
        public ObservableCollection<ImageItem> StylesPreviewList { get; set; }

        // UI model - list that holds styles variations items
        public ObservableCollection<ImageItem> StylesVariationList { get; set; }

        // a background presentation that will do the preview processing
        public StylesHandler PreviewPresentation { get; set; }

        // indicate whether the window is open/closed or not
        public bool IsOpen { get; set; }
        public bool IsClosing { get; set; }

        // used to refresh preview and variation images
        private DateTime _latestPreviewUpdateTime = DateTime.Now;
        private DateTime _latestImageChangedTime = DateTime.Now;

        // used to indicate right-click item
        private int _rightClickedSearchListBoxItemIndex = -1;

        // used to clean up unused image files
        private readonly HashSet<string> _imageFilesInUse = new HashSet<string>();

        private bool _isWindowActivatedWithPreview = true;
        private bool _isStylePreviewRegionInit;

        # endregion

        #region Initialization

        public ImagesLabWindow()
        {
            InitializeComponent();
            InitImageSelectionList();
            InitStylesPreviewList();
            InitStylesVariationList();
            InitFontFamilyList();
            IsOpen = true;
            if (TempPath.InitTempFolder() && StoragePath.InitPersistentFolder(_imageFilesInUse))
            {
                InitGotoSlideDialog();
                InitReloadStylesDialog();
                InitPreviewPresentation();
                InitDragAndDrop();
            }
            else
            {
                ShowErrorMessageBox(TextCollection.ImagesLabText.ErrorFailToInitTempFolder);
            }
        }

        private void InitStylesVariationList()
        {
            StylesVariationList = new ObservableCollection<ImageItem>();
            StylesVariationList.CollectionChanged += StylesVariationList_OnCollectionChanged;
            StylesVariationListBox.DataContext = this;
        }

        private void InitStylesPreviewList()
        {
            StylesPreviewList = new ObservableCollection<ImageItem>();
            StylesPreviewList.CollectionChanged += StylesPreviewList_OnCollectionChanged;
            StylesPreviewListBox.DataContext = this;
        }

        private void InitPreviewPresentation()
        {
            PreviewPresentation = new StylesHandler();
            PreviewPresentation.Open(withWindow: false, focus: false);
        }

        # endregion

        # region Common UI Events & Interactions

        private void ImageSelectionList_OnCollectionChanged(object sender, NotifyCollectionChangedEventArgs e)
        {
            ImageSelectionInstructions.Visibility = ImageSelectionList.Count == 0
                ? Visibility.Visible
                : Visibility.Hidden;
            if (ImageSelectionInstructions.Visibility == Visibility.Visible)
            {
                PreviewInstructions.Visibility = Visibility.Hidden;
                PreviewInstructionsWhenNoSelectedSlide.Visibility = Visibility.Hidden;
            }
            else
            {
                PreviewInstructions.Visibility = Visibility.Visible;
                PreviewInstructionsWhenNoSelectedSlide.Visibility = Visibility.Hidden;
            }

            // show StylesPreviewRegion aft there'r some images in the SearchList region
            if (ImageSelectionList.Count > 0 && !_isStylePreviewRegionInit)
            {
                // only one entry
                _isStylePreviewRegionInit = true;
                StylesPreviewGrid.Visibility = Visibility.Visible;
            }
        }

        private void StylesVariationList_OnCollectionChanged(object sender, NotifyCollectionChangedEventArgs e)
        {
            if (StylesVariationList.Count != 0)
            {
                VariationInstructionsWhenNoSelectedSlide.Visibility = Visibility.Hidden;
                VariantsComboBox.IsEnabled = true;
                VariantsColorPanel.IsEnabled = true;
            }
            else if (PowerPointCurrentPresentationInfo.CurrentSlide == null)
            {
                VariationInstructionsWhenNoSelectedSlide.Visibility = Visibility.Visible;
                VariantsComboBox.IsEnabled = false;
                VariantsColorPanel.IsEnabled = false;
            }
            else // select 'loading' image
            {
                VariationInstructionsWhenNoSelectedSlide.Visibility = Visibility.Hidden;
                VariantsComboBox.IsEnabled = false;
                VariantsColorPanel.IsEnabled = false;
            }
        }

        private void StylesPreviewList_OnCollectionChanged(object sender, NotifyCollectionChangedEventArgs e)
        {
            PreviewInstructions.BeginAnimation(OpacityProperty, null);
            PreviewInstructions.Opacity = 1f;
            PreviewInstructionsWhenNoSelectedSlide.BeginAnimation(OpacityProperty, null);
            PreviewInstructionsWhenNoSelectedSlide.Opacity = 1f;

            if (StylesPreviewList.Count != 0 || ImageSelectionInstructions.Visibility == Visibility.Visible)
            {
                PreviewInstructions.Visibility = Visibility.Hidden;
                PreviewInstructionsWhenNoSelectedSlide.Visibility = Visibility.Hidden;
            }
            else if (PowerPointCurrentPresentationInfo.CurrentSlide == null)
            {
                PreviewInstructionsWhenNoSelectedSlide.Visibility = Visibility.Visible;
                PreviewInstructions.Visibility = Visibility.Hidden;
            }
            else
            {
                PreviewInstructions.Visibility = Visibility.Visible;
                PreviewInstructionsWhenNoSelectedSlide.Visibility = Visibility.Hidden;
            }
        }

        private void ImageSelectButton_OnClick(object sender, RoutedEventArgs e)
        {
            DoLoadImageFromFile();
        }

        // intent:
        // do previewing, when search result item is (not) selected
        private void ImageSelectionListBox_OnSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            var source = (ImageItem) ImageSelectionListBox.SelectedValue;
            if (source == null)
            {
                CloseFlyouts();
            }
            _latestImageChangedTime = DateTime.Now;
            DoPreview();
        }

        // intent:
        // allow arrow keys to navigate the search result items in the list
        private void ListBox_OnKeyDown(object sender, KeyEventArgs e)
        {
            var listbox = sender as ListBox;
            if (listbox == null || listbox.Items.Count <= 0)
            {
                return;
            }

            switch (e.Key)
            {
                case Key.Right:
                case Key.Down:
                    if (!listbox.Items.MoveCurrentToNext())
                    {
                        listbox.Items.MoveCurrentToLast();
                    }
                    break;

                case Key.Left:
                case Key.Up:
                    if (!listbox.Items.MoveCurrentToPrevious())
                    {
                        listbox.Items.MoveCurrentToFirst();
                    }
                    break;

                default:
                    return;
            }

            e.Handled = true;
            var item = (ListBoxItem) listbox.ItemContainerGenerator
                .ContainerFromItem(listbox.SelectedItem);
            item.Focus();
        }

        private void ImageSelectionListBox_OnKeyDown(object sender, KeyEventArgs e)
        {
            switch (e.Key)
            {
                case Key.Delete:
                case Key.Back:
                    DeleteSelectedImageShape();
                    return;
            }
            ListBox_OnKeyDown(sender, e);
        }

        // intent: drag splitter to change grid width
        private void Splitter_OnDragDelta(object sender, DragDeltaEventArgs e)
        {
            // need remove animation before set its width
            StylesPreviewGrid.BeginAnimation(WidthProperty, null);
            StylesPreviewGrid.Width = StylesPreviewGrid.ActualWidth + e.HorizontalChange;
        }

        // enable & disable insert button
        private void StylesPreivewListBox_OnSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (StylesPreviewListBox.SelectedValue != null)
            {
                StylesPickUpButton.IsEnabled = true;
                StylesApplyButton.IsEnabled = true;
            }
            else
            {
                StylesPickUpButton.IsEnabled = false;
                StylesApplyButton.IsEnabled = false;
            }
        }

        private void StylesPreviewListBox_OnKeyUp(object sender, KeyEventArgs e)
        {
            var listbox = sender as ListBox;
            if (listbox == null || listbox.Items.Count <= 0)
            {
                return;
            }

            switch (e.Key)
            {
                case Key.Enter:
                    if (StylesApplyButton.IsEnabled)
                    {
                        StylesApplyButton.RaiseEvent(new RoutedEventArgs(ButtonBase.ClickEvent));
                    }
                    break;
            }
        }

        private void ImagesLabWindow_OnActivated(object sender, EventArgs e)
        {
            if (!_isWindowActivatedWithPreview) return;

            if (PowerPointCurrentPresentationInfo.CurrentSlide == null)
            {
                GotoSlideButton.IsEnabled = false;
                ReloadStylesButton.IsEnabled = false;
            }
            else
            {
                GotoSlideButton.IsEnabled = true;
                ReloadStylesButton.IsEnabled = true;
            }

            if (QuickDropDialog != null && QuickDropDialog.IsOpen)
            {
                QuickDropDialog.Hide();
                QuickDropDialog.IsOpen = false;
            }

            Dispatcher.BeginInvoke(new Action(() =>
            {
                if (_isVariationsFlyoutOpen)
                {
                    UpdateStyleVariationsImages();
                }
                else
                {
                    DoPreview();
                }
            }));
        }

        private void ImageSelectionListBox_OnPreviewMouseDown(object sender, MouseButtonEventArgs e)
        {
            var item = ItemsControl.ContainerFromElement((ItemsControl) sender, (DependencyObject) e.OriginalSource) as ListBoxItem;
            if (item == null || item.Content == null) return;

            // intent: clicking 'load more' should not change selection
            if (e.RightButton == MouseButtonState.Pressed)
            {
                _rightClickedSearchListBoxItemIndex = -1;
                for (var i = 0; i < ImageSelectionListBox.Items.Count; i++)
                {
                    var lbi = ImageSelectionListBox.ItemContainerGenerator.ContainerFromIndex(i) as ListBoxItem;
                    if (lbi == null) continue;
                    if (IsMouseOverTarget(lbi, e.GetPosition(lbi)))
                    {
                        _rightClickedSearchListBoxItemIndex = i;
                        break;
                    }
                }
                e.Handled = true;
            }
        }

        private static bool IsMouseOverTarget(Visual target, Point point)
        {
            var bounds = VisualTreeHelper.GetDescendantBounds(target);
            return bounds.Contains(point);
        }

        private void MenuItemDeleteThisImage_OnClick(object sender, RoutedEventArgs e)
        {
            DeleteImageShape();
        }

        private void MenuItemDeleteAllImages_OnClick(object sender, RoutedEventArgs e)
        {
            DeleteAllImageShapes();
        }

        private void DeleteAllImageShapes()
        {
            ImageSelectionList.Clear();
        }

        private void DeleteImageShape()
        {
            if (_rightClickedSearchListBoxItemIndex < 0 
                || _rightClickedSearchListBoxItemIndex > ImageSelectionListBox.Items.Count)
                return;

            var selectedImage = (ImageItem) ImageSelectionListBox.Items.GetItemAt(_rightClickedSearchListBoxItemIndex);
            if (selectedImage == null) return;

            ImageSelectionList.RemoveAt(_rightClickedSearchListBoxItemIndex);
        }

        private void DeleteSelectedImageShape()
        {
            var selectedImage = (ImageItem)ImageSelectionListBox.SelectedItem;
            if (selectedImage == null) return;

            ImageSelectionList.RemoveAt(ImageSelectionListBox.SelectedIndex);
        }

        # endregion
    }
}
