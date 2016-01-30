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
using System.Windows.Media.Animation;
using PowerPointLabs.Models;
using PowerPointLabs.PictureSlidesLab.Model;
using PowerPointLabs.PictureSlidesLab.Util;
using PowerPointLabs.PictureSlidesLab.View.Interface;
using PowerPointLabs.PictureSlidesLab.ViewModel;
using PowerPointLabs.WPF.Observable;
using ButtonBase = System.Windows.Controls.Primitives.ButtonBase;
using DragEventArgs = System.Windows.DragEventArgs;
using KeyEventArgs = System.Windows.Input.KeyEventArgs;
using ListBox = System.Windows.Controls.ListBox;

namespace PowerPointLabs.PictureSlidesLab.View
{
    /// <summary>
    /// Interaction logic for Picture Slides Lab
    /// </summary>
    public partial class PictureSlidesLabWindow : IPictureSlidesLabWindowView
    {
        # region Props & States
        // View model that contains the presenting logic
        private PictureSlidesLabWindowViewModel ViewModel { set; get; }

        // UI model used to adjust image offset
        public AdjustImageWindow CropWindow { get; set; }

        // UI model for drag and drop
        public ObservableString DragAndDropInstructionText { get; set; }
        public QuickDropDialog QuickDropDialog { get; set; }

        // indicate to add-in that whether the window is open
        public bool IsOpen { get; set; }

        // indicate that whether the window is closing
        private bool IsClosing { get; set; }

        public bool IsVariationsFlyoutOpen { get; private set; }

        // used to indicate right-click item
        private int _clickedImageSelectionItemIndex = -1;

        // other UI control flags
        private bool _isAbleLoadingOnWindowActivate = true;
        private bool _isStylePreviewRegionInit;

        private bool _isFirstTimeOnActivated = true;

        # endregion

        #region Lifecycle

        public PictureSlidesLabWindow()
        {
            InitializeComponent();

            InitViewModel();
            InitGotoSlideDialog();
            InitLoadStylesDialog();
            InitDragAndDrop();
            IsOpen = true;
        }

        private void InitViewModel()
        {
            ViewModel = new PictureSlidesLabWindowViewModel(this);
            ViewModel.StylesVariationList.CollectionChanged += StylesVariationList_OnCollectionChanged;
            ViewModel.StylesPreviewList.CollectionChanged += StylesPreviewList_OnCollectionChanged;
            ViewModel.ImageSelectionList.CollectionChanged += ImageSelectionList_OnCollectionChanged;
            DataContext = ViewModel;

            UpdatePreviewInterfaceWhenImageListChange(ViewModel.ImageSelectionList);
        }

        private void InitDragAndDrop()
        {
            DragAndDropInstructionText = new ObservableString { Text = TextCollection.PictureSlidesLabText.DragAndDropInstruction };
            DragAndDropInstructions.DataContext = DragAndDropInstructionText;
        }

        private void PictureSlidesLabWindow_OnClosing(object sender, CancelEventArgs e)
        {
            IsOpen = false;
            IsClosing = true;
            if (QuickDropDialog != null)
            {
                QuickDropDialog.Close();
            }
            ViewModel.CleanUp();
        }

        # endregion

        # region Common UI Events & Interactions

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

        #region Drag and Drop
        private void PictureSlidesLabWindow_OnDragLeave(object sender, DragEventArgs args)
        {
            PictureSlidesLabGridOverlay.Visibility = Visibility.Hidden;
        }

        private void PictureSlidesLabWindow_OnDragEnter(object sender, DragEventArgs args)
        {
            if (args.Data.GetDataPresent("FileDrop")
                || args.Data.GetDataPresent("Text"))
            {
                PictureSlidesLabGridOverlay.Visibility = Visibility.Visible;
                DisableLoadingStyleOnWindowActivate();
                Activate();
                EnableLoadingStyleOnWindowActivate();
            }
        }

        private void PictureSlidesLabWindow_OnDrop(object sender, DragEventArgs args)
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
                    ViewModel.AddImageSelectionListItem(imageUrl, 
                        PowerPointCurrentPresentationInfo.CurrentSlide.GetNativeSlide(),
                        PowerPointPresentation.Current.SlideWidth,
                        PowerPointPresentation.Current.SlideHeight);
                }
            }
            finally
            {
                PictureSlidesLabGridOverlay.Visibility = Visibility.Hidden;
            }
        }
        #endregion

        /// <summary>
        /// Show QuickDrop dialog when PictureSlidesLab window is deactivated
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void PictureSlidesLabWindow_OnDeactivated(object sender, EventArgs e)
        {
            if (!IsClosing
                && (CropWindow == null || !CropWindow.IsOpen)
                && (QuickDropDialog == null || !QuickDropDialog.IsOpen))
            {
                QuickDropDialog = new QuickDropDialog(this);
                QuickDropDialog.DropHandler += PictureSlidesLabWindow_OnDrop;
                QuickDropDialog.Show();
            }
        }

        /// <summary>
        /// Show preview images when an image is selected in the ImageSelectionList
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ImageSelectionListBox_OnSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (ImageSelectionListBox.SelectedValue != null)
            {
                EnableUpdatingPreviewImages();
            }
            if (ImageSelectionListBox.SelectedIndex == 0)
            {
                // 0 is for Choose Pictures placeholder,
                // de-select it
                ImageSelectionListBox.SelectedIndex = -1;
            }
            UpdatePreviewImages();
        }

        private void StylesCustomizeButton_OnClick(object sender, RoutedEventArgs e)
        {
            CustomizeStyle(
                (ImageItem) ImageSelectionListBox.SelectedValue ?? CreateDefaultPictureItem());
        }

        private void StylesPreviewApplyButton_OnClick(object sender, RoutedEventArgs e)
        {
            ViewModel.ApplyStyleInPreviewStage(
                PowerPointCurrentPresentationInfo.CurrentSlide.GetNativeSlide(),
                PowerPointPresentation.Current.SlideWidth,
                PowerPointPresentation.Current.SlideHeight);
        }

        /// <summary>
        /// Allow arrow keys to navigate the listbox items
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
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

        /// <summary>
        /// Delete image by backspace/delete key
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ImageSelectionListBox_OnKeyDown(object sender, KeyEventArgs e)
        {
            switch (e.Key)
            {
                case Key.Delete:
                case Key.Back:
                    DeleteSelectedImage();
                    return;
            }
            ListBox_OnKeyDown(sender, e);
        }

        /// <summary>
        /// Drag splitter to change grid width
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Splitter_OnDragDelta(object sender, DragDeltaEventArgs e)
        {
            StylesPreviewGrid.Width = StylesPreviewGrid.ActualWidth + e.HorizontalChange;
        }

        /// <summary>
        /// Enable & disable Apply button for preview interface
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void StylesPreivewListBox_OnSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            UpdatePreviewStageControls();
        }

        private void UpdatePreviewStageControls()
        {
            if (StylesPreviewListBox.SelectedValue != null
                && _isDisplayDefaultPicture)
            {
                StylesCustomizeButton.IsEnabled = true;
                StylesApplyButton.IsEnabled = false;
            }
            else if (StylesPreviewListBox.SelectedValue != null
                     && ImageSelectionListBox.SelectedValue != null)
            {
                StylesCustomizeButton.IsEnabled = true;
                StylesApplyButton.IsEnabled = true;
            }
            else
            {
                StylesCustomizeButton.IsEnabled = false;
                StylesApplyButton.IsEnabled = false;
            }
        }

        /// <summary>
        /// Press ENTER button to apply
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
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

        /// <summary>
        /// When window is re-activated, refresh the preview images and hide QuickDrop dialog
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void PictureSlidesLabWindow_OnActivated(object sender, EventArgs e)
        {
            if (QuickDropDialog != null && QuickDropDialog.IsOpen)
            {
                QuickDropDialog.Hide();
                QuickDropDialog.IsOpen = false;
            }

            if (PowerPointCurrentPresentationInfo.CurrentSlide == null)
            {
                GotoSlideButton.IsEnabled = false;
                LoadStylesButton.IsEnabled = false;
                ViewModel.StylesPreviewList.Clear();
                ViewModel.StylesVariationList.Clear();
            }
            else if (_isStylePreviewRegionInit && _isAbleLoadingOnWindowActivate)
            {
                GotoSlideButton.IsEnabled = true;
                LoadStylesButton.IsEnabled = true;
                Dispatcher.BeginInvoke(new Action(() =>
                {
                    if (_isFirstTimeOnActivated)
                    {
                        _isFirstTimeOnActivated = false;
                    }
                    else
                    {
                        LoadStyleAndImage(PowerPointCurrentPresentationInfo.CurrentSlide);
                    }
                }));
            }
        }

        /// <summary>
        /// Obtain right-clicked listbox item and don't select any image
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
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
            else if (e.LeftButton == MouseButtonState.Pressed)
            {
                var imageItem = item.Content as ImageItem;
                if (imageItem != null
                    && imageItem.ImageFile == StoragePath.ChoosePicturesImgPath)
                {
                    var openFileDialog = new OpenFileDialog
                    {
                        Multiselect = true,
                        Filter = @"Image File|*.png;*.jpg;*.jpeg;*.bmp;*.gif;"
                    };
                    DisableLoadingStyleOnWindowActivate();
                    if (openFileDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                    {
                        ViewModel.AddImageSelectionListItem(openFileDialog.FileNames);
                    }
                    EnableLoadingStyleOnWindowActivate();
                    e.Handled = true;
                }
            }
        }

        private void MenuItemDeleteThisImage_OnClick(object sender, RoutedEventArgs e)
        {
            DeleteImage();
        }

        private void MenuItemDeleteAllImages_OnClick(object sender, RoutedEventArgs e)
        {
            DeleteAllImage();
        }

        private void MenuItemAdjustImage_OnClick(object sender, RoutedEventArgs e)
        {

            if (_clickedImageSelectionItemIndex < 0
                || _clickedImageSelectionItemIndex > ImageSelectionListBox.Items.Count)
                return;

            var selectedImage = (ImageItem)ImageSelectionListBox.Items.GetItemAt(_clickedImageSelectionItemIndex);
            if (selectedImage == null || selectedImage.ImageFile == StoragePath.LoadingImgPath) return;

            AdjustImageDimensions(selectedImage);
        }

        private void MenuItemAdjustImage_OnClickFromPreviewListBox(object sender, RoutedEventArgs e)
        {
            var selectedImage = (ImageItem) ImageSelectionListBox.SelectedItem;
            if (selectedImage == null || selectedImage.ImageFile == StoragePath.LoadingImgPath) return;

            AdjustImageDimensions(selectedImage);
        }

        /// <summary>
        /// Update controls states when selection changed in the variation stage
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void VariationListBox_OnSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (ImageSelectionListBox.SelectedValue != null
                && StylesVariationListBox.SelectedValue != null
                && StylesPreviewListBox.SelectedValue != null)
            {
                StyleApplyButton.IsEnabled = true;
                ViewModel.UpdateStyleVariationStyleOptionsWhenSelectedItemChange();
                UpdateVariationStageControls();
            }
            else if (_isDisplayDefaultPicture
                     && StylesVariationListBox.SelectedValue != null
                     && StylesPreviewListBox.SelectedValue != null)
            {
                StyleApplyButton.IsEnabled = false;
                ViewModel.UpdateStyleVariationStyleOptionsWhenSelectedItemChange();
                UpdateVariationStageControls();
            }
            else
            {
                StyleApplyButton.IsEnabled = false;
            }
        }

        /// <summary>
        /// step-by-step customization when user changes variant category
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void VariantsComboBox_OnSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            ViewModel.UpdateStepByStepStylesVariationImages(
                (ImageItem) ImageSelectionListBox.SelectedValue ?? CreateDefaultPictureItem(),
                PowerPointCurrentPresentationInfo.CurrentSlide.GetNativeSlide(),
                PowerPointPresentation.Current.SlideWidth,
                PowerPointPresentation.Current.SlideHeight);
        }

        private void StylesVariationApplyButton_OnClick(object sender, RoutedEventArgs e)
        {
            ViewModel.ApplyStyleInVariationStage(
                PowerPointCurrentPresentationInfo.CurrentSlide.GetNativeSlide(),
                PowerPointPresentation.Current.SlideWidth,
                PowerPointPresentation.Current.SlideHeight);
        }

        private void VariationFlyoutBackButton_OnClick(object sender, RoutedEventArgs e)
        {
            CloseVariationsFlyout();
        }

        #endregion

        #region Helper funcs

        private void DeleteAllImage()
        {
            ViewModel.RemoveAllImageSelectionListItems();
        }

        private void DeleteImage()
        {
            if (_clickedImageSelectionItemIndex < 1 
                || _clickedImageSelectionItemIndex >= ImageSelectionListBox.Items.Count)
                return;

            var selectedImage = (ImageItem) ImageSelectionListBox.Items.GetItemAt(_clickedImageSelectionItemIndex);
            if (selectedImage == null) return;

            ViewModel.ImageSelectionList.RemoveAt(_clickedImageSelectionItemIndex);
        }

        private void DeleteSelectedImage()
        {
            var selectedImage = (ImageItem)ImageSelectionListBox.SelectedItem;
            if (selectedImage == null
                || ImageSelectionListBox.SelectedIndex == 0)
                return;

            ViewModel.ImageSelectionList.RemoveAt(ImageSelectionListBox.SelectedIndex);
        }

        private void AdjustImageDimensions(ImageItem source)
        {
            CropWindow = new AdjustImageWindow();
            CropWindow.SetThumbnailImage(source.ImageFile);
            CropWindow.SetFullsizeImage(source.FullSizeImageFile);
            if (source.Rect.Width > 1)
            {
                CropWindow.SetCropRect(source.Rect.X, source.Rect.Y, source.Rect.Width, source.Rect.Height);
            }

            DisableLoadingStyleOnWindowActivate();
            CropWindow.ShowAdjustPictureDimensionsDialog();
            EnableLoadingStyleOnWindowActivate();

            if (CropWindow.IsCropped)
            {
                source.UpdateImageAdjustmentOffset(CropWindow.CropResult, CropWindow.CropResultThumbnail, CropWindow.Rect);
                UpdatePreviewImages();
            }
        }

        /// <summary>
        /// decide visibility for instructions and stylesPreviewGrid
        /// </summary>
        private void UpdatePreviewInterfaceWhenImageListChange(Collection<ImageItem> list)
        {
            // there is only `Choose Picture` placeholder image
            ImageSelectionInstructions.Visibility = list.Count <= 1
                ? Visibility.Visible
                : Visibility.Hidden;

            if (ImageSelectionInstructions.Visibility == Visibility.Visible)
            {
                PreviewInstructions.Visibility = Visibility.Hidden;
                PreviewInstructionsWhenNoSelectedSlide.Visibility = Visibility.Hidden;
                VariationInstructions.Visibility = Visibility.Hidden;
                VariationInstructionsWhenNoSelectedSlide.Visibility = Visibility.Hidden;
            }
            else if (ImageSelectionListBox.SelectedValue == null)
            {
                PreviewInstructions.Visibility = Visibility.Visible;
                PreviewInstructionsWhenNoSelectedSlide.Visibility = Visibility.Hidden;
                VariationInstructions.Visibility = Visibility.Visible;
                VariationInstructionsWhenNoSelectedSlide.Visibility = Visibility.Hidden;
            }

            // there are `Choose Picture` placeholder image +
            // 2 sample pictures + maybe 1 image chosen by user
            if (StoragePath.IsFirstTimeUsage() && list.Count <= 4
                && ImageSelectionInstructions.Visibility == Visibility.Hidden)
            {
                PreviewInstructions.Visibility = Visibility.Hidden;
                ImageSelectionInstructionsForFirstTime.Visibility = Visibility.Visible;
            }
            else
            {
                ImageSelectionInstructionsForFirstTime.Visibility = Visibility.Hidden;
            }

            // show StylesPreviewRegion aft there'r some images in the SearchList region
            if (list.Count > 1 && !_isStylePreviewRegionInit)
            {
                // only one time entry
                _isStylePreviewRegionInit = true;
                StylesPreviewGrid.Visibility = Visibility.Visible;
                GotoSlideButton.IsEnabled = true;
                LoadStylesButton.IsEnabled = true;
            }
        }

        /// <summary>
        /// decide visibility and enability of variation stage's 
        /// </summary>
        private void UpdateVariationInterface(Collection<ImageItem> list)
        {
            if (list.Count != 0)
            {
                VariationInstructions.Visibility = Visibility.Hidden;
                VariationInstructionsWhenNoSelectedSlide.Visibility = Visibility.Hidden;
                VariantsComboBox.IsEnabled = true;
                VariantsColorPanel.IsEnabled = true;
            }
            else if (PowerPointCurrentPresentationInfo.CurrentSlide == null)
            {
                VariationInstructions.Visibility = Visibility.Hidden;
                VariationInstructionsWhenNoSelectedSlide.Visibility = Visibility.Visible;
                VariantsComboBox.IsEnabled = false;
                VariantsColorPanel.IsEnabled = false;
            }
            else // select 'loading' image
            {
                VariationInstructions.Visibility = Visibility.Visible;
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

        private void OpenVariationsFlyout()
        {
            Dispatcher.BeginInvoke(new Action(() =>
            {
                if (IsVariationsFlyoutOpen) return;

                var left2RightToShowTranslate = new TranslateTransform { X = -StylesPreviewGrid.ActualWidth };
                StyleVariationsFlyout.RenderTransform = left2RightToShowTranslate;
                StyleVariationsFlyout.Visibility = Visibility.Visible;
                var left2RightToShowAnimation = new DoubleAnimation(-StylesPreviewGrid.ActualWidth, 0,
                    TimeSpan.FromMilliseconds(350))
                {
                    EasingFunction = new SineEase { EasingMode = EasingMode.EaseInOut },
                    AccelerationRatio = 0.5
                };

                left2RightToShowTranslate.BeginAnimation(TranslateTransform.XProperty, left2RightToShowAnimation);
                IsVariationsFlyoutOpen = true;
            }));
        }

        private void CloseVariationsFlyout()
        {
            if (!IsVariationsFlyoutOpen) return;

            var right2LeftToHideTranslate = new TranslateTransform();
            StyleVariationsFlyout.RenderTransform = right2LeftToHideTranslate;
            var right2LeftToHideAnimation = new DoubleAnimation(0, -StyleVariationsFlyout.ActualWidth,
                TimeSpan.FromMilliseconds(350))
            {
                EasingFunction = new SineEase { EasingMode = EasingMode.EaseInOut },
                AccelerationRatio = 0.5
            };
            right2LeftToHideAnimation.Completed += (sender, args) =>
            {
                Dispatcher.BeginInvoke(new Action(() =>
                {
                    StyleVariationsFlyout.Visibility = Visibility.Collapsed;
                    if (IsEnableUpdatingPreviewImages())
                    {
                        UpdatePreviewImages();
                    }
                    else
                    {
                        EnableUpdatingPreviewImages();
                        UpdatePreviewImages(CreateDefaultPictureItem());
                        DisableUpdatingPreviewImages();
                        UpdatePreviewStageControls();
                    }
                }));
            };

            right2LeftToHideTranslate.BeginAnimation(TranslateTransform.XProperty, right2LeftToHideAnimation);
            IsVariationsFlyoutOpen = false;
        }

        private void UpdatePreviewImages(ImageItem source = null)
        {
            if (!IsEnableUpdatingPreviewImages()) return;

            ViewModel.UpdatePreviewImages(
                source ?? (ImageItem) ImageSelectionListBox.SelectedValue,
                PowerPointCurrentPresentationInfo.CurrentSlide.GetNativeSlide(),
                PowerPointPresentation.Current.SlideWidth,
                PowerPointPresentation.Current.SlideHeight);
        }

        private void CustomizeStyle(ImageItem source, List<StyleOptions> givenStyles = null,
            Dictionary<string, List<StyleVariants>> givenVariants = null)
        {
            ViewModel.UpdateStyleVariationImagesWhenOpenFlyout(
                source ?? (ImageItem) ImageSelectionListBox.SelectedValue,
                PowerPointCurrentPresentationInfo.CurrentSlide.GetNativeSlide(),
                PowerPointPresentation.Current.SlideWidth,
                PowerPointPresentation.Current.SlideHeight,
                givenStyles, givenVariants);
            OpenVariationsFlyout();
        }

        private void EnterDefaultPictureMode()
        {
            DisableUpdatingPreviewImages();
        }

        private bool IsEnableUpdatingPreviewImages()
        {
            return !_isDisplayDefaultPicture;
        }

        private void EnableLoadingStyleOnWindowActivate()
        {
            _isAbleLoadingOnWindowActivate = true;
        }

        private void DisableLoadingStyleOnWindowActivate()
        {
            _isAbleLoadingOnWindowActivate = false;
        }

        #endregion
    }
}
