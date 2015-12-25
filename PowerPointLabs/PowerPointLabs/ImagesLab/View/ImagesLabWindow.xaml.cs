using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Collections.Specialized;
using System.ComponentModel;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Forms;
using System.Windows.Input;
using System.Windows.Media;
using PowerPointLabs.ImagesLab.Model;
using PowerPointLabs.ImagesLab.Util;
using PowerPointLabs.ImagesLab.View.Interface;
using PowerPointLabs.ImagesLab.ViewModel;
using PowerPointLabs.Models;
using PowerPointLabs.WPF.Observable;
using ButtonBase = System.Windows.Controls.Primitives.ButtonBase;
using DragEventArgs = System.Windows.DragEventArgs;
using KeyEventArgs = System.Windows.Input.KeyEventArgs;
using ListBox = System.Windows.Controls.ListBox;

namespace PowerPointLabs.ImagesLab.View
{
    /// <summary>
    /// Interaction logic for Images Lab
    /// </summary>
    public partial class ImagesLabWindow : IImagesLabWindow
    {
        # region Props & States
        // view model that contains the presenting logic
        private ImagesLabWindowViewModel ViewModel { set; get; }

        // used to adjust image offset
        public AdjustImageWindow CropWindow { get; set; }

        // UI model for drag and drop
        public ObservableString DragAndDropInstructionText { get; set; }
        public QuickDropDialog QuickDropDialog { get; set; }

        // indicate to add-in that whether the window is open
        public bool IsOpen { get; set; }

        // indicate that whether the window is closing
        private bool IsClosing { get; set; }

        // used to refresh preview and variation images
        private DateTime _latestPreviewUpdateTime = DateTime.Now;
        private DateTime _latestImageChangedTime = DateTime.Now;

        // used to indicate right-click item
        private int _clickedImageSelectionItemIndex = -1;

        // other control flags
        private bool _isWindowActivatedWithPreview = true;
        private bool _isStylePreviewRegionInit;
        private bool _isVariationsFlyoutOpen;

        // list that holds font families
        private readonly List<string> _fontFamilyList = new List<string>();

        # endregion

        #region Initialization

        public ImagesLabWindow()
        {
            InitializeComponent();

            InitViewModel();

            InitFontFamilyList();
            InitGotoSlideDialog();
            InitReloadStylesDialog();
            InitDragAndDrop();

            IsOpen = true;
        }

        private void InitViewModel()
        {
            ViewModel = new ImagesLabWindowViewModel(this);
            ViewModel.StylesVariationList.CollectionChanged += StylesVariationList_OnCollectionChanged;
            ViewModel.StylesPreviewList.CollectionChanged += StylesPreviewList_OnCollectionChanged;
            ViewModel.ImageSelectionList.CollectionChanged += ImageSelectionList_OnCollectionChanged;
            DataContext = ViewModel;
            UpdatePreviewInterfaceWhenImageListChange(ViewModel.ImageSelectionList);
        }

        private void InitDragAndDrop()
        {
            // TODO move to text collection
            DragAndDropInstructionText = new ObservableString { Text = "Drag and Drop here to get image." };
            DragAndDropInstructions.DataContext = DragAndDropInstructionText;
        }

        # endregion

        # region Common UI Events & Interactions

        private void ImagesLabWindow_OnClosing(object sender, CancelEventArgs e)
        {
            IsOpen = false;
            IsClosing = true;
            if (QuickDropDialog != null)
            {
                QuickDropDialog.Close();
            }
            ViewModel.CleanUp();
        }

        private void ImageSelectionList_OnCollectionChanged(object sender, NotifyCollectionChangedEventArgs e)
        {
            UpdatePreviewInterfaceWhenImageListChange(sender as Collection<ImageItem>);
        }

        private void StylesVariationList_OnCollectionChanged(object sender, NotifyCollectionChangedEventArgs e)
        {
            UpdateVariationInterface(sender as Collection<ImageItem>);
        }

        private void StylesPreviewList_OnCollectionChanged(object sender, NotifyCollectionChangedEventArgs e)
        {
            UpdatePreviewInterfaceWhenPreviewListChange(sender as Collection<ImageItem>);
        }

        private void ImageSelectButton_OnClick(object sender, RoutedEventArgs e)
        {
            var openFileDialog = new OpenFileDialog
            {
                Multiselect = true,
                Filter = @"Image File|*.png;*.jpg;*.jpeg;*.bmp;*.gif;"
            };
            if (openFileDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                ViewModel.AddImageSelectionListItem(openFileDialog.FileNames);
            }
        }

        private void ImagesLabWindow_OnDragLeave(object sender, DragEventArgs args)
        {
            ImagesLabGridOverlay.Visibility = Visibility.Hidden;
        }

        private void ImagesLabWindow_OnDragEnter(object sender, DragEventArgs args)
        {
            if (args.Data.GetDataPresent("FileDrop")
                || args.Data.GetDataPresent("Text"))
            {
                ImagesLabGridOverlay.Visibility = Visibility.Visible;
                _isWindowActivatedWithPreview = false;
                Activate();
                _isWindowActivatedWithPreview = true;
            }
        }

        private void ImagesLabWindow_OnDrop(object sender, DragEventArgs args)
        {
            try
            {
                if (args == null) return;

                if (args.Data.GetDataPresent("FileDrop"))
                {
                    var filenames = (args.Data.GetData("FileDrop") as string[]);
                    if (filenames == null || filenames.Length == 0) return;

                    ViewModel.AddImageSelectionListItem(filenames);
                }
                else if (args.Data.GetDataPresent("Text"))
                {
                    var imageUrl = args.Data.GetData("Text") as string;
                    ViewModel.AddImageSelectionListItem(imageUrl);
                }
            }
            finally
            {
                ImagesLabGridOverlay.Visibility = Visibility.Hidden;
            }
        }

        private void ImagesLabWindow_OnDeactivated(object sender, EventArgs e)
        {
            if (!IsClosing
                && (CropWindow == null || !CropWindow.IsOpen)
                && (QuickDropDialog == null || !QuickDropDialog.IsOpen))
            {
                QuickDropDialog = new QuickDropDialog(this);
                QuickDropDialog.DropHandler += ImagesLabWindow_OnDrop;
                QuickDropDialog.Show();
            }
        }

        private void ImageSelectionListBox_OnSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            var source = (ImageItem) ImageSelectionListBox.SelectedValue;
            if (source == null && _isVariationsFlyoutOpen)
            {
                CloseVariationsFlyout();
            }
            _latestImageChangedTime = DateTime.Now;
            UpdatePreviewImages();
        }

        private void StylesPickUpButton_OnClick(object sender, RoutedEventArgs e)
        {
            CustomizeStyle();
        }

        private void StylesApplyButton_OnClick(object sender, RoutedEventArgs e)
        {
            if (StylesPreviewListBox.SelectedValue == null) return;

            var source = ImageSelectionListBox.SelectedValue as ImageItem;
            var targetStyle = ((ImageItem)StylesPreviewListBox.SelectedValue).Tooltip;
            ViewModel.ApplyStyleInPreviewStage(source, targetStyle);
        }

        // intent:
        // allow arrow keys to navigate the listbox items
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

        // intent:
        // delete image by backspace/delete key
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

        // intent: 
        // drag splitter to change grid width
        private void Splitter_OnDragDelta(object sender, DragDeltaEventArgs e)
        {
            StylesPreviewGrid.Width = StylesPreviewGrid.ActualWidth + e.HorizontalChange;
        }

        // intent:
        // enable & disable Apply button for preview interface
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

        // intent:
        // press ENTER button to apply
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

            UpdatePreviewImages();
        }

        // intent:
        // obtain right-clicked listbox item
        private void ImageSelectionListBox_OnPreviewMouseDown(object sender, MouseButtonEventArgs e)
        {
            var item = ItemsControl.ContainerFromElement((ItemsControl) sender, (DependencyObject) e.OriginalSource) 
                as ListBoxItem;
            if (item == null || item.Content == null) return;

            if (e.RightButton == MouseButtonState.Pressed)
            {
                _clickedImageSelectionItemIndex = -1;
                for (var i = 0; i < ImageSelectionListBox.Items.Count; i++)
                {
                    var listBoxItem = ImageSelectionListBox.ItemContainerGenerator.ContainerFromIndex(i) 
                        as ListBoxItem;
                    if (listBoxItem == null) continue;

                    if (IsMouseOverTarget(listBoxItem, e.GetPosition(listBoxItem)))
                    {
                        _clickedImageSelectionItemIndex = i;
                        break;
                    }
                }
                e.Handled = true;
            }
        }

        private void MenuItemDeleteThisImage_OnClick(object sender, RoutedEventArgs e)
        {
            DeleteImageShape();
        }

        private void MenuItemDeleteAllImages_OnClick(object sender, RoutedEventArgs e)
        {
            DeleteAllImageShapes();
        }

        private void MenuItemAdjustImage_OnClick(object sender, RoutedEventArgs e)
        {

            if (_clickedImageSelectionItemIndex < 0
                || _clickedImageSelectionItemIndex > ImageSelectionListBox.Items.Count)
                return;

            var selectedImage = (ImageItem)ImageSelectionListBox.Items.GetItemAt(_clickedImageSelectionItemIndex);
            if (selectedImage == null || selectedImage.ImageFile == StoragePath.LoadingImgPath) return;

            AdjustImageOffset(selectedImage);
        }

        private void MenuItemAdjustImage_OnClickFromPreviewListBox(object sender, RoutedEventArgs e)
        {
            var selectedImage = (ImageItem) ImageSelectionListBox.SelectedItem;
            if (selectedImage == null || selectedImage.ImageFile == StoragePath.LoadingImgPath) return;

            AdjustImageOffset(selectedImage);
        }

        #endregion

        #region Helper funcs

        private void UpdatePreviewImages()
        {
            var image = (ImageItem)ImageSelectionListBox.SelectedValue;
            if (image == null || image.ImageFile == StoragePath.LoadingImgPath)
            {
                if (_isVariationsFlyoutOpen)
                {
                    ViewModel.ClearStyleVariationList();
                }
                else
                {
                    ViewModel.ClearStylesPreviewList();
                }
            }
            else if (_isVariationsFlyoutOpen)
            {
                UpdateStyleVariationsImages();
            }
            else
            {
                var selectedId = StylesPreviewListBox.SelectedIndex;
                ViewModel.UpdatePreviewImages(image);

                StylesPreviewListBox.SelectedIndex = selectedId;
                _latestPreviewUpdateTime = DateTime.Now;
            }
        }

        private void DeleteAllImageShapes()
        {
            ViewModel.ClearImageSelectionList();
        }

        private void DeleteImageShape()
        {
            if (_clickedImageSelectionItemIndex < 0 
                || _clickedImageSelectionItemIndex > ImageSelectionListBox.Items.Count)
                return;

            var selectedImage = (ImageItem) ImageSelectionListBox.Items.GetItemAt(_clickedImageSelectionItemIndex);
            if (selectedImage == null) return;

            ViewModel.RemoveImageSelectionListItem(_clickedImageSelectionItemIndex);
        }

        private void DeleteSelectedImageShape()
        {
            var selectedImage = (ImageItem)ImageSelectionListBox.SelectedItem;
            if (selectedImage == null) return;

            ViewModel.RemoveImageSelectionListItem(ImageSelectionListBox.SelectedIndex);
        }

        private void AdjustImageOffset(ImageItem source)
        {
            CropWindow = new AdjustImageWindow();
            CropWindow.SetThumbnailImage(source.ImageFile);
            CropWindow.SetFullsizeImage(source.FullSizeImageFile);
            if (source.Rect.Width > 1)
            {
                CropWindow.SetCropRect(source.Rect.X, source.Rect.Y, source.Rect.Width, source.Rect.Height);
            }
            CropWindow.IsOpen = true;
            CropWindow.ShowDialog();
            CropWindow.IsOpen = false;

            if (CropWindow.IsCropped)
            {
                source.UpdateImageAdjustmentOffset(CropWindow.CropResult, CropWindow.CropResultThumbnail, CropWindow.Rect);
            }
        }

        /// <summary>
        /// decide visibility for instructions and stylesPreviewGrid
        /// </summary>
        private void UpdatePreviewInterfaceWhenImageListChange(Collection<ImageItem> list)
        {
            ImageSelectionInstructions.Visibility = list.Count == 0
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
            if (list.Count > 0 && !_isStylePreviewRegionInit)
            {
                // only one time entry
                _isStylePreviewRegionInit = true;
                StylesPreviewGrid.Visibility = Visibility.Visible;
            }
        }

        /// <summary>
        /// decide visibility and enability of variation stage's 
        /// </summary>
        private void UpdateVariationInterface(Collection<ImageItem> list)
        {
            if (list.Count != 0)
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

        /// <summary>
        /// decide visibility for instructions
        /// </summary>
        private void UpdatePreviewInterfaceWhenPreviewListChange(Collection<ImageItem> list)
        {
            if (list.Count != 0 || ImageSelectionInstructions.Visibility == Visibility.Visible)
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

        private static bool IsMouseOverTarget(Visual target, Point point)
        {
            var bounds = VisualTreeHelper.GetDescendantBounds(target);
            return bounds.Contains(point);
        }
        #endregion
    }
}
