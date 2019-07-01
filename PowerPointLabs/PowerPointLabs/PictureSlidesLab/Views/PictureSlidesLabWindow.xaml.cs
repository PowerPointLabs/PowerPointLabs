using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Collections.Specialized;
using System.ComponentModel;
using System.Drawing;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Animation;
using System.Windows.Threading;

using MahApps.Metro.Controls.Dialogs;

using Microsoft.Office.Interop.PowerPoint;

using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.ActionFramework.Common.Log;
using PowerPointLabs.PictureSlidesLab.Model;
using PowerPointLabs.PictureSlidesLab.Service;
using PowerPointLabs.PictureSlidesLab.Util;
using PowerPointLabs.PictureSlidesLab.ViewModel;
using PowerPointLabs.PictureSlidesLab.Views.Interface;
using PowerPointLabs.TextCollection;
using PowerPointLabs.Utils;
using PowerPointLabs.Utils.Windows;
using PowerPointLabs.WPF.Observable;

using ButtonBase = System.Windows.Controls.Primitives.ButtonBase;
using Clipboard = System.Windows.Forms.Clipboard;
using DragEventArgs = System.Windows.DragEventArgs;
using KeyEventArgs = System.Windows.Input.KeyEventArgs;
using ListBox = System.Windows.Controls.ListBox;
using MessageBox = PowerPointLabs.Utils.Windows.MessageBoxUtil;
using Point = System.Windows.Point;

namespace PowerPointLabs.PictureSlidesLab.Views
{
    /// <summary>
    /// Interaction logic for Picture Slides Lab
    /// </summary>
    public partial class PictureSlidesLabWindow : IPictureSlidesLabWindowView
    {
        # region Props & States
        // UI model used to adjust image offset
        public AdjustImageWindow CropWindow { get; set; }

        // UI model for drag and drop
        public ObservableString DragAndDropInstructionText { get; set; }
        public QuickDropDialog QuickDropDialog { get; set; }

        // indicate to add-in that whether the window is open
        public bool IsOpen { get; set; }

        public bool IsVariationsFlyoutOpen { get; private set; }

        // indicate that whether the window is closing
        private bool IsClosing { get; set; }

        // View model that contains the presenting logic
        private PictureSlidesLabWindowViewModel ViewModel { set; get; }

        // used to indicate right-click item
        private int _clickedImageSelectionItemIndex = -1;

        // other UI control flags
        private bool _isAbleLoadingOnWindowActivate = true;
        private bool _isStylePreviewRegionInit;
        private int _lastSelectedSlideIndex = -1;
        private bool _isDisplayDefaultPicture;
        private bool _isEnableUpdatePreview = true;

        //Window size control constant
        private const double StandardSystemWidth = 1280.0;
        private const double StandardSystemHeight = 800.0;
        private const double StandardWindowWidth = 1200.0;
        private const double StandardWindowHeight = 700.0;
        private const double StandardPrewienGridWidth = 560.0;
        private const float StandardDpi = 96f;

        # endregion

        #region Lifecycle

        public PictureSlidesLabWindow()
        {
            InitializeComponent();
            // start loading process
            EveryDayPhrase.Text = new EveryDayPhraseService().GetEveryDayPhrase();
            PictureSlidesLabGridLoadingOverlay.Visibility = Visibility.Visible;
            IsOpen = true;
            SettingsButtonIcon.Source = GraphicsUtil.BitmapToImageSource(Properties.Resources.PslSettings);
            PictureAspectRefreshButtonIcon.Source = GraphicsUtil.BitmapToImageSource(Properties.Resources.PslRefresh);
            InitSizePosition();
            Logger.Log("PSL begins");

            SetTimeout(Init, 800);
        }

        private void Init()
        {
            try
            {
                InitUiExceptionHandling();
                InitViewModel();
                InitGotoSlideDialog();
                InitLoadStylesDialog();
                InitErrorTextDialog();
                InitDragAndDrop();
                // leave some time for data binding to finish
                SetTimeout(InitStyleing, 50);
                Logger.Log("PSL init done");
            }
            catch (Exception e)
            {
                ShowErrorMessageBox(PictureSlidesLabText.ErrorWhenInitialize, e);
                Logger.LogException(e, "Init");
            }
        }

        private void InitSizePosition()
        {
            System.Drawing.Size mSize = WinformUtil.WorkingAreaSize;
            //Devices might have scale factors > 100%
            float scaleFactor = GetScalingFactor();
            double systemHeight = mSize.Height / scaleFactor;
            double systemWidth = mSize.Width / scaleFactor;
            double windowWidth = systemWidth / StandardSystemWidth;
            double windowHeight = systemHeight / StandardSystemHeight;

            this.Window.Width = StandardWindowWidth * windowWidth;
            this.Window.Height = StandardWindowHeight * windowHeight;
            this.Window.StylesPreviewGrid.Width = StandardPrewienGridWidth * windowWidth;
            this.Window.Left = (systemWidth - this.Window.Width) / 2;
            this.Window.Top = (systemHeight - this.Window.Height) / 2;
            this.Window.WindowStartupLocation = WindowStartupLocation.Manual;
        }
        
        private float GetScalingFactor()
        {
            Graphics graphics = Graphics.FromHwnd(IntPtr.Zero);
            return graphics.DpiX / StandardDpi;
        }

        private void InitUiExceptionHandling()
        {
            AppDomain.CurrentDomain.UnhandledException += HandleUnhandledException;
            Dispatcher.UnhandledException += HandleUnhandledException;
            Logger.Log("PSL init UI exception handling done");
        }

        private void InitStyleing()
        {
            try
            {
                Logger.Log("PSL init styling begins");
                // load back the style from the current slide, or
                // select the first picture to preview styles
                bool isSuccessfullyLoaded = LoadStyleAndImage(this.GetCurrentSlide(),
                    isLoadingWithDefaultPicture: false);
                if (ViewModel.ImageSelectionList.Count >= 2 && !isSuccessfullyLoaded)
                {
                    Logger.Log("Not loaded back style and picture, going to select a picture.");
                    // index-0 is choosePicture placeholder
                    ViewModel.ImageSelectionListSelectedId.Number = 1;
                }
                Logger.Log("PSL init styling done");
            }
            catch (Exception e)
            {
                ShowErrorMessageBox(PictureSlidesLabText.ErrorWhenInitialize, e);
                Logger.LogException(e, "InitStyleing");
            }
            finally
            {
                // remove loading overlay
                PictureSlidesLabGridLoadingOverlay.Visibility = Visibility.Collapsed;
                Logger.Log("PSL init loading screen collapsed");
            }
        }

        private void InitViewModel()
        {
            ViewModel = new PictureSlidesLabWindowViewModel(this);
            ViewModel.StylesVariationList.CollectionChanged += StylesVariationList_OnCollectionChanged;
            ViewModel.StylesPreviewList.CollectionChanged += StylesPreviewList_OnCollectionChanged;
            ViewModel.ImageSelectionList.CollectionChanged += ImageSelectionList_OnCollectionChanged;
            DataContext = ViewModel;
            SettingsPane.DataContext = ViewModel.Settings;
            SettingsPane.UpdateInsertCitationControlsVisibility();

            UpdatePreviewInterfaceWhenImageListChange(ViewModel.ImageSelectionList);
            Logger.Log("PSL init ViewModel done");
        }

        private void InitDragAndDrop()
        {
            DragAndDropInstructionText = new ObservableString { Text = PictureSlidesLabText.DragAndDropInstruction };
            DragAndDropInstructions.DataContext = DragAndDropInstructionText;
            Logger.Log("PSL init drag and drop done");
        }

        private void PictureSlidesLabWindow_OnClosing(object sender, CancelEventArgs e)
        {
            try
            {
                IsOpen = false;
                IsClosing = true;
                if (QuickDropDialog != null)
                {
                    QuickDropDialog.Close();
                    QuickDropDialog = null;
                }
                ViewModel.CleanUp();
                Logger.Log("PSL closed");
            }
            catch (Exception expt)
            {
                Logger.LogException(expt, "PictureSlidesLabWindow_OnClosing");
            }
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
                Logger.Log("Drag enter");
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
                if (args == null)
                {
                    return;
                }

                Logger.Log("Drop enter");
                if (args.Data.GetDataPresent("FileDrop"))
                {
                    string[] filenames = (args.Data.GetData("FileDrop") as string[]);
                    if (filenames == null || filenames.Length == 0)
                    {
                        return;
                    }

                    ViewModel.AddImageSelectionListItem(filenames,
                        this.GetCurrentSlide().GetNativeSlide(),
                        this.GetCurrentPresentation().SlideWidth,
                        this.GetCurrentPresentation().SlideHeight);
                }
                else if (args.Data.GetDataPresent("Text"))
                {
                    string imageUrl = args.Data.GetData("Text") as string;
                    ViewModel.AddImageSelectionListItem(imageUrl,
                        this.GetCurrentSlide().GetNativeSlide(),
                        this.GetCurrentPresentation().SlideWidth,
                        this.GetCurrentPresentation().SlideHeight);
                }
            }
            catch (Exception e)
            {
                Logger.LogException(e, "PictureSlidesLabWindow_OnDrop");
            }
            finally
            {
                PictureSlidesLabGridOverlay.Visibility = Visibility.Hidden;
            }
        }
        #endregion

        #region Copy and Paste Picture

        private void MenuItemPastePictureHere_OnClick(object sender, RoutedEventArgs e)
        {
            HandlePastedPicture();
        }

        private void PictureSlidesLabWindow_OnKeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.V
                && Keyboard.Modifiers == ModifierKeys.Control)
            {
                HandlePastedPicture();
            }
        }

        /// <summary>
        /// set isUsingWinformMsgBox to true when it requires a msgbox out of
        /// the main window of PSL
        /// </summary>
        /// <param name="isUsingWinformMsgBox"></param>
        private void HandlePastedPicture(bool isUsingWinformMsgBox = false)
        {
            try
            {
                System.Drawing.Image pastedPicture = Clipboard.GetImage();
                StringCollection pastedFiles = Clipboard.GetFileDropList();

                if (pastedPicture == null &&
                    (pastedFiles == null || pastedFiles.Count == 0))
                {
                    if (isUsingWinformMsgBox)
                    {
                        MessageBox.Show(PictureSlidesLabText.InfoPasteNothing, "PowerPointLabs");
                    }
                    else
                    {
                        ShowInfoMessageBox(PictureSlidesLabText.InfoPasteNothing);
                    }
                    Logger.Log("Nothing to paste");
                    return;
                }

                if (pastedPicture != null)
                {
                    Logger.Log("Pasted enter");
                    string pastedPictureFile = StoragePath.GetPath("pastedImg-"
                                                                + DateTime.Now.GetHashCode() + "-"
                                                                + Guid.NewGuid().ToString().Substring(0, 7));
                    pastedPicture.Save(pastedPictureFile);
                    ViewModel.AddImageSelectionListItem(new[] {pastedPictureFile},
                        this.GetCurrentSlide().GetNativeSlide(),
                        this.GetCurrentPresentation().SlideWidth,
                        this.GetCurrentPresentation().SlideHeight);

                    // examine whether it's thumbnail picture
                    if (pastedPicture.Width <= 400
                        && pastedPicture.Height <= 400)
                    {
                        if (isUsingWinformMsgBox)
                        {
                            MessageBox.Show(PictureSlidesLabText.InfoPasteThumbnail, "PowerPointLabs");
                        }
                        else
                        {
                            ShowInfoMessageBox(PictureSlidesLabText.InfoPasteThumbnail);
                        }
                    }
                }
                else if (pastedFiles != null && pastedFiles.Count > 0)
                {
                    ViewModel.AddImageSelectionListItem(pastedFiles.Cast<string>().ToArray(),
                        this.GetCurrentSlide().GetNativeSlide(),
                        this.GetCurrentPresentation().SlideWidth,
                        this.GetCurrentPresentation().SlideHeight);
                }
            }
            catch (Exception e)
            {
                Logger.LogException(e, "HandlePastedPicture");
            }
        }

        #endregion

        private void PictureSlidesLabGridLoadingOverlay_OnPreviewMouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            PictureSlidesLabGridLoadingOverlay.Visibility = Visibility.Collapsed;
        }

        /// <summary>
        /// Show QuickDrop dialog when PictureSlidesLab window is deactivated
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void PictureSlidesLabWindow_OnDeactivated(object sender, EventArgs e)
        {
            try
            {
                if (IsDisposed)
                {
                    return;
                }

                _lastSelectedSlideIndex = this.GetCurrentSlide().Index;

                if (!IsDisposed
                    && (CropWindow == null || !CropWindow.IsOpen))
                {
                    if (QuickDropDialog == null)
                    {
                        QuickDropDialog = new QuickDropDialog(this);
                        QuickDropDialog.DropHandler += PictureSlidesLabWindow_OnDrop;
                        QuickDropDialog.PasteHandler += () => { HandlePastedPicture(isUsingWinformMsgBox: true); };
                        QuickDropDialog.ShowQuickDropDialog();
                        Logger.Log("PSL Quick Drop Dialog begins");
                    }
                    else if (!QuickDropDialog.IsOpen && !QuickDropDialog.IsDisposed)
                    {
                        QuickDropDialog.ShowQuickDropDialog();
                    }
                }
            }
            catch (Exception expt)
            {
                Logger.LogException(expt, "PictureSlidesLabWindow_OnDeactivated");
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
                LeaveDefaultPictureMode();
            }
            if (ImageSelectionListBox.SelectedIndex == 0)
            {
                // 0 is for Choose Pictures placeholder,
                // de-select it
                ImageSelectionListBox.SelectedIndex = -1;
            }
            ViewModel.UpdateSelectedPictureInPictureVariation();
            UpdatePreviewImages();
        }

        private void StylesCustomizeButton_OnClick(object sender, RoutedEventArgs e)
        {
            CustomizeStyle(
                (ImageItem) ImageSelectionListBox.SelectedValue ?? CreateDefaultPictureItem());
        }

        private void StylesPreviewApplyButton_OnClick(object sender, RoutedEventArgs e)
        {
            this.StartNewUndoEntry();
            ViewModel.ApplyStyleInPreviewStage(
                this.GetCurrentSlide().GetNativeSlide(),
                this.GetCurrentPresentation().SlideWidth,
                this.GetCurrentPresentation().SlideHeight);
            GC.Collect();
        }

        /// <summary>
        /// Allow arrow keys to navigate the listbox items
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ListBox_OnKeyDown(object sender, KeyEventArgs e)
        {
            ListBox listbox = sender as ListBox;
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
            double newWidth = StylesPreviewGrid.ActualWidth + e.HorizontalChange;
            // Prevent StylesPreviewGrid from becoming too small
            if (newWidth < StylesPreviewGrid.MinWidth)
            {
                StylesPreviewGrid.Width = StylesPreviewGrid.MinWidth;
            }
            // Prevent StylesPreviewGrid from overflowing grid column 0
            else if (newWidth > Column0.Width.Value)
            {
                StylesPreviewGrid.Width = Column0.Width.Value;
            }
            else
            {
                StylesPreviewGrid.Width = newWidth;
            }
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
                && IsDisplayDefaultPicture())
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
            ListBox listbox = sender as ListBox;
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
            try
            {
                // init last selected slide index
                if (_lastSelectedSlideIndex == -1)
                {
                    _lastSelectedSlideIndex = this.GetCurrentSlide().Index;
                }

                // hide quick drop dialog when main window activated
                if (QuickDropDialog != null && QuickDropDialog.IsOpen)
                {
                    QuickDropDialog.HideQuickDropDialog();
                }

                // when no current slide
                if (this.GetCurrentSlide() == null)
                {
                    Logger.Log("Current slide is null");
                    GotoSlideButton.IsEnabled = false;
                    LoadStylesButton.IsEnabled = false;
                    ViewModel.StylesPreviewList.Clear();
                    ViewModel.StylesVariationList.Clear();
                }
                // when allowed to do loading
                else if (_isStylePreviewRegionInit && _isAbleLoadingOnWindowActivate)
                {
                    GotoSlideButton.IsEnabled = true;
                    LoadStylesButton.IsEnabled = true;
                    SetTimeout(() =>
                    {
                        // update preview images when slide no change
                        if (_lastSelectedSlideIndex == this.GetCurrentSlide().Index)
                        {
                            UpdatePreviewImages();
                        }
                        // or load style and image if slide has been changed
                        else
                        {
                            LoadStyleAndImage(this.GetCurrentSlide());
                        }
                    }, 250);
                }
            }
            catch (Exception expt)
            {
                Logger.LogException(expt, "PictureSlidesLabWindow_OnActivated");
            }
        }

        /// <summary>
        /// Obtain right-clicked listbox item and don't select any image
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ImageSelectionListBox_OnPreviewMouseDown(object sender, MouseButtonEventArgs e)
        {
            ListBoxItem item = ItemsControl.ContainerFromElement((ItemsControl) sender, (DependencyObject) e.OriginalSource) 
                as ListBoxItem;
            if (item == null || item.Content == null)
            {
                return;
            }

            if (e.RightButton == MouseButtonState.Pressed)
            {
                _clickedImageSelectionItemIndex = -1;
                for (int i = 0; i < ImageSelectionListBox.Items.Count; i++)
                {
                    ListBoxItem listBoxItem = ImageSelectionListBox.ItemContainerGenerator.ContainerFromIndex(i) 
                        as ListBoxItem;
                    if (listBoxItem == null)
                    {
                        continue;
                    }

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
                Logger.Log("begin import pictures");
                ImageItem imageItem = item.Content as ImageItem;
                if (imageItem != null
                    && imageItem.ImageFile == StoragePath.ChoosePicturesImgPath)
                {
                    DisableLoadingStyleOnWindowActivate();
                    List<string> openedFiles = OpenFileDialogUtil.MultiOpen(
                        filter: @"Image File|*.png;*.jpg;*.jpeg;*.bmp;*.gif;");

                    if (openedFiles != null)
                    {
                        ViewModel.AddImageSelectionListItem(openedFiles.ToArray(),
                            this.GetCurrentSlide().GetNativeSlide(),
                            this.GetCurrentPresentation().SlideWidth,
                            this.GetCurrentPresentation().SlideHeight);
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
            ShowInfoMessageBox(PictureSlidesLabText.InfoDeleteAllImage, 
                MessageDialogStyle.AffirmativeAndNegative)
                .ContinueWith(task =>
                {
                    if (task.Result == MessageDialogResult.Affirmative)
                    {
                        Dispatcher.BeginInvoke(new Action(DeleteAllImage));
                    }
                });
        }

        private void MenuItemAdjustImage_OnClick(object sender, RoutedEventArgs e)
        {
            if (_clickedImageSelectionItemIndex < 0
                || _clickedImageSelectionItemIndex > ImageSelectionListBox.Items.Count)
            {
                return;
            }

            ImageItem selectedImage = (ImageItem)ImageSelectionListBox.Items.GetItemAt(_clickedImageSelectionItemIndex);
            if (selectedImage == null || selectedImage.ImageFile == StoragePath.LoadingImgPath)
            {
                return;
            }

            AdjustImageDimensions(selectedImage);
        }

        private void MenuItemAdjustImage_OnClickFromPreviewListBox(object sender, RoutedEventArgs e)
        {
            if (ViewModel.IsInPictureVariation())
            {
                ImageItem imageItem = ViewModel.GetSelectedPictureInPictureVariation(
                    StylesVariationListBox.SelectedIndex);
                if (imageItem.ImageFile == StoragePath.NoPicturePlaceholderImgPath
                    || imageItem.ImageFile == StoragePath.LoadingImgPath)
                {
                    return;
                }
                AdjustImageDimensions(imageItem);
            }
            else
            {
                ImageItem selectedImage = (ImageItem)ImageSelectionListBox.SelectedItem;
                if (selectedImage == null || selectedImage.ImageFile == StoragePath.LoadingImgPath)
                {
                    return;
                }

                AdjustImageDimensions(selectedImage);
            }
        }

        private void MenuItemEditSource_OnClick(object sender, RoutedEventArgs e)
        {
            if (_clickedImageSelectionItemIndex < 0
                || _clickedImageSelectionItemIndex > ImageSelectionListBox.Items.Count)
            {
                return;
            }

            ImageItem selectedImage = (ImageItem)ImageSelectionListBox.Items.GetItemAt(_clickedImageSelectionItemIndex);
            if (selectedImage == null || selectedImage.ImageFile == StoragePath.LoadingImgPath)
            {
                return;
            }

            EditPictureSource(selectedImage);
        }

        private void MenuItemEditSource_OnClickFromPreviewListBox(object sender, RoutedEventArgs e)
        {
            if (ViewModel.IsInPictureVariation())
            {
                ImageItem imageItem = ViewModel.GetSelectedPictureInPictureVariation(
                    StylesVariationListBox.SelectedIndex);
                if (imageItem.ImageFile == StoragePath.NoPicturePlaceholderImgPath
                    || imageItem.ImageFile == StoragePath.LoadingImgPath)
                {
                    return;
                }
                EditPictureSource(imageItem);
            }
            else
            {
                ImageItem selectedImage = (ImageItem)ImageSelectionListBox.SelectedItem;
                if (selectedImage == null || selectedImage.ImageFile == StoragePath.LoadingImgPath)
                {
                    return;
                }

                EditPictureSource(selectedImage);
            }
        }

        /// <summary>
        /// Update controls states when selection changed in the variation stage
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void VariationListBox_OnSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (ViewModel.IsInPictureVariation()
                     && StylesVariationListBox.SelectedIndex >= 0)
            {
                ImageItem selectedImageItem =
                    ViewModel
                    .GetSelectedPictureInPictureVariation(StylesVariationListBox.SelectedIndex);
                if (selectedImageItem.ImageFile == StoragePath.NoPicturePlaceholderImgPath
                    || selectedImageItem.ImageFile == StoragePath.LoadingImgPath)
                {
                    StyleVariationApplyButton.IsEnabled = false;
                }
                else
                {
                    StyleVariationApplyButton.IsEnabled = true;
                }
                ViewModel.UpdateStyleVariationStyleOptionsWhenSelectedItemChange();
                UpdateVariationStageControls();
            }
            else if (ImageSelectionListBox.SelectedValue != null
                && StylesVariationListBox.SelectedValue != null
                && StylesPreviewListBox.SelectedValue != null)
            {
                StyleVariationApplyButton.IsEnabled = true;
                ViewModel.UpdateStyleVariationStyleOptionsWhenSelectedItemChange();
                UpdateVariationStageControls();
            }
            else if (IsDisplayDefaultPicture()
                     && StylesVariationListBox.SelectedValue != null
                     && StylesPreviewListBox.SelectedValue != null)
            {
                StyleVariationApplyButton.IsEnabled = false;
                ViewModel.UpdateStyleVariationStyleOptionsWhenSelectedItemChange();
                UpdateVariationStageControls();
            }
            else
            {
                StyleVariationApplyButton.IsEnabled = false;
            }
        }

        /// <summary>
        /// step-by-step customization when user changes variant category
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void VariantsComboBox_OnSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            try
            {
                Logger.Log("Change aspect to " + ViewModel.CurrentVariantCategory.Text);
                ViewModel.UpdateStepByStepStylesVariationImages(
                    (ImageItem) ImageSelectionListBox.SelectedValue ?? CreateDefaultPictureItem(),
                    this.GetCurrentSlide().GetNativeSlide(),
                    this.GetCurrentPresentation().SlideWidth,
                    this.GetCurrentPresentation().SlideHeight);
            }
            catch (Exception expt)
            {
                ShowErrorMessageBox(TextCollection.PictureSlidesLabText.ErrorFailedToLoadPreviewImages, expt);
                Logger.LogException(expt, "VariantsComboBox_OnSelectionChanged");
            }
        }

        private void StylesVariationApplyButton_OnClick(object sender, RoutedEventArgs e)
        {
            this.StartNewUndoEntry();
            ViewModel.ApplyStyleInVariationStage(
                this.GetCurrentSlide().GetNativeSlide(),
                this.GetCurrentPresentation().SlideWidth,
                this.GetCurrentPresentation().SlideHeight);
            GC.Collect();
        }

        private void VariationFlyoutBackButton_OnClick(object sender, RoutedEventArgs e)
        {
            CloseVariationsFlyout();
        }

        private void AddCitationSlideButton_OnClick(object sender, RoutedEventArgs e)
        {
            Models.PowerPointPresentation presentation = this.GetCurrentPresentation();
            if (presentation.Slides.Any(PictureCitationSlide.IsCitationSlide))
            {
                Models.PowerPointSlide citationSlide = presentation.Slides.Where(PictureCitationSlide.IsCitationSlide).First();
                ViewModel.AddPictureCitationSlide(citationSlide.GetNativeSlide(), presentation.Slides);
            }
            else // no citation slide yet, so create one
            {
                Slide slide = presentation.Presentation.Slides.Add(presentation.SlideCount + 1, PpSlideLayout.ppLayoutText);
                ViewModel.AddPictureCitationSlide(slide, presentation.Slides);
            }
            presentation.AddAckSlide();
            ShowInfoMessageBox(PictureSlidesLabText.InfoAddPictureCitationSlide);
        }

        private void OpenSettingsButton_OnClick(object sender, RoutedEventArgs e)
        {
            SettingsFlyoutControl.IsOpen = true;
        }

        private void SettingsFlyoutControl_OnIsOpenChanged(object sender, RoutedEventArgs e)
        {
            if (!SettingsFlyoutControl.IsOpen)
            {
                Logger.Log("Setting flyout is closed");
                SetTimeout(() => { UpdatePreviewImages(); }, 500);
            }
        }

        private void PictureAspectRefreshButton_OnClick(object sender, RoutedEventArgs e)
        {
            ViewModel.RefreshLast8Pictures();
            UpdatePreviewImages();
        }

        #endregion

        #region Helper funcs

        private void DeleteAllImage()
        {
            ViewModel.RemoveAllImageSelectionListItems();
            HandleDeletePictureInPictureVariation();
        }

        private void HandleDeletePictureInPictureVariation()
        {
            ViewModel.UpdatePictureInPictureVariationWhenDeleteSome();
            if (ViewModel.IsInPictureVariation())
            {
                UpdatePreviewImages();
            }
        }

        private void DeleteImage()
        {
            if (_clickedImageSelectionItemIndex < 1 
                || _clickedImageSelectionItemIndex >= ImageSelectionListBox.Items.Count)
            {
                return;
            }

            ImageItem selectedImage = (ImageItem) ImageSelectionListBox.Items.GetItemAt(_clickedImageSelectionItemIndex);
            if (selectedImage == null)
            {
                return;
            }

            ViewModel.ImageSelectionList.RemoveAt(_clickedImageSelectionItemIndex);
            HandleDeletePictureInPictureVariation();
        }

        private void DeleteSelectedImage()
        {
            ImageItem selectedImage = (ImageItem)ImageSelectionListBox.SelectedItem;
            if (selectedImage == null
                || ImageSelectionListBox.SelectedIndex == 0)
            {
                return;
            }

            ViewModel.ImageSelectionList.RemoveAt(ImageSelectionListBox.SelectedIndex);
            HandleDeletePictureInPictureVariation();
        }

        private void EditPictureSource(ImageItem source)
        {
            MetroDialogSettings metroDialogSettings = new MetroDialogSettings
            {
                DefaultText = source.Source
            };
            this.ShowInputAsync("Edit Picture Source", "Picture taken from", metroDialogSettings)
                .ContinueWith(task =>
                {
                    if (!string.IsNullOrEmpty(task.Result))
                    {
                        source.Source = task.Result;
                    }
                });
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
                if (CropWindow.IsRotated)
                {
                    source.Source = CropWindow.RotateResult;
                    source.ImageFile = ImageUtil.GetThumbnailFromFullSizeImg(CropWindow.RotateResult);
                    source.FullSizeImageFile = CropWindow.RotateResult;
                    source.ContextLink = CropWindow.RotateResult;
                }
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
            else if (ImageSelectionListBox.SelectedValue == null
                     && StylesPreviewListBox.Items.Count == 0
                     && StylesVariationListBox.Items.Count == 0)
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
                VariantsSlider.IsEnabled = true;
            }
            else if (this.GetCurrentSlide() == null)
            {
                VariationInstructions.Visibility = Visibility.Hidden;
                VariationInstructionsWhenNoSelectedSlide.Visibility = Visibility.Visible;
                VariantsComboBox.IsEnabled = false;
                VariantsColorPanel.IsEnabled = false;
                VariantsSlider.IsEnabled = false;
            }
            else // select 'loading' image
            {
                VariationInstructions.Visibility = Visibility.Visible;
                VariationInstructionsWhenNoSelectedSlide.Visibility = Visibility.Hidden;
                VariantsComboBox.IsEnabled = false;
                VariantsColorPanel.IsEnabled = false;
                VariantsSlider.IsEnabled = false;
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
            else if (this.GetCurrentSlide() == null)
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
            Rect bounds = VisualTreeHelper.GetDescendantBounds(target);
            return bounds.Contains(point);
        }

        private void OpenVariationsFlyout()
        {
            Dispatcher.BeginInvoke(new Action(() =>
            {
                if (IsVariationsFlyoutOpen)
                {
                    return;
                }

                TranslateTransform left2RightToShowTranslate = new TranslateTransform { X = -StylesPreviewGrid.ActualWidth };
                StyleVariationsFlyout.RenderTransform = left2RightToShowTranslate;
                StyleVariationsFlyout.Visibility = Visibility.Visible;
                DoubleAnimation left2RightToShowAnimation = new DoubleAnimation(-StylesPreviewGrid.ActualWidth, 0,
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
            if (!IsVariationsFlyoutOpen)
            {
                return;
            }

            TranslateTransform right2LeftToHideTranslate = new TranslateTransform();
            StyleVariationsFlyout.RenderTransform = right2LeftToHideTranslate;
            DoubleAnimation right2LeftToHideAnimation = new DoubleAnimation(0, -StyleVariationsFlyout.ActualWidth,
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
                    if (!IsDisplayDefaultPicture())
                    {
                        UpdatePreviewImages();
                    }
                    else
                    {
                        UpdatePreviewImages(CreateDefaultPictureItem());
                        UpdatePreviewStageControls();
                    }
                }));
            };

            right2LeftToHideTranslate.BeginAnimation(TranslateTransform.XProperty, right2LeftToHideAnimation);
            IsVariationsFlyoutOpen = false;
            ViewModel.CurrentVariantCategory.Text = "";
            Logger.Log("Variation is closed");
        }

        private void UpdatePreviewImages(ImageItem source = null, bool isEnteringPictureVariation = false)
        {
            try
            {
                if (!_isEnableUpdatePreview)
                {
                    return;
                }

                if (IsVariationsFlyoutOpen && isEnteringPictureVariation)
                {
                    Logger.Log("Entering pic aspect");
                    // when it's to load the design for a default picture,
                    // and it's at the variation stage,
                    // directly jump to picture variation to select picture
                    int picVariationIndex = ViewModel.VariantsCategory.IndexOf(
                        PictureSlidesLabText.VariantCategoryPicture);
                    if (VariantsComboBox.SelectedIndex != picVariationIndex)
                    {
                        VariantsComboBox.SelectedIndex = picVariationIndex;
                    }
                    else
                    {
                        ViewModel.UpdatePreviewImages(
                        source ?? (ImageItem)ImageSelectionListBox.SelectedValue ?? CreateDefaultPictureItem(),
                        this.GetCurrentSlide().GetNativeSlide(),
                        this.GetCurrentPresentation().SlideWidth,
                        this.GetCurrentPresentation().SlideHeight);
                    }
                }
                else if (IsDisplayDefaultPicture())
                {
                    Logger.Log("In default pic mode");
                    // if it's in Default Picture mode, allow
                    // updating preview images
                    ViewModel.UpdatePreviewImages(
                        source ?? CreateDefaultPictureItem(),
                        this.GetCurrentSlide().GetNativeSlide(),
                        this.GetCurrentPresentation().SlideWidth,
                        this.GetCurrentPresentation().SlideHeight);
                }
                else
                {
                    // else, try to update preview images using
                    // source or current selected picture.
                    ViewModel.UpdatePreviewImages(
                        source ?? (ImageItem) ImageSelectionListBox.SelectedValue,
                        this.GetCurrentSlide().GetNativeSlide(),
                        this.GetCurrentPresentation().SlideWidth,
                        this.GetCurrentPresentation().SlideHeight);
                }
            }
            catch (Exception e)
            {
                Logger.LogException(e, "UpdatePreviewImages");
            }
        }

        private void CustomizeStyle(ImageItem source, List<StyleOption> givenStyles = null,
            Dictionary<string, List<StyleVariant>> givenVariants = null)
        {
            try
            {
                ViewModel.UpdateStyleVariationImagesWhenOpenFlyout(
                    source ?? (ImageItem) ImageSelectionListBox.SelectedValue,
                    this.GetCurrentSlide().GetNativeSlide(),
                    this.GetCurrentPresentation().SlideWidth,
                    this.GetCurrentPresentation().SlideHeight,
                    givenStyles, givenVariants);
                OpenVariationsFlyout();
            }
            catch (Exception e)
            {
                Logger.LogException(e, "CustomizeStyle");
            }
        }

        private void LeaveDefaultPictureMode()
        {
            _isDisplayDefaultPicture = false;
        }

        private void EnableLoadingStyleOnWindowActivate()
        {
            _isAbleLoadingOnWindowActivate = true;
        }

        private void DisableLoadingStyleOnWindowActivate()
        {
            _isAbleLoadingOnWindowActivate = false;
        }

        /// <summary>
        /// Execute action after time (in ms)
        /// </summary>
        /// <param name="action"></param>
        /// <param name="time">time in ms</param>
        private void SetTimeout(Action action, int time)
        {
            DispatcherTimer timer = new DispatcherTimer(DispatcherPriority.Render)
            {
                Interval = new TimeSpan(0, 0, 0, 0, time) // in ms
            };
            timer.Tick += (o, args) =>
            {
                timer.Stop();
                action();
            };
            timer.Start();
        }

        private void HandleUnhandledException(object sender, DispatcherUnhandledExceptionEventArgs e)
        {
            Logger.LogException(e.Exception, sender.GetType() + " " + sender);
            ShowErrorMessageBox("Unexpected errors happened!", e.Exception);
        }

        private void HandleUnhandledException(object sender, UnhandledExceptionEventArgs e)
        {
            Logger.LogException(e.ExceptionObject as Exception, sender.GetType() + " " + sender);
            ShowErrorMessageBox("Unexpected errors happened!", e.ExceptionObject as Exception);
        }

        // check PSL window is really closing or not
        private bool IsDisposed
        {
            get
            {
                if (IsClosing)
                {
                    return true;
                }

                System.Reflection.PropertyInfo propertyInfo = typeof(Window).GetProperty("IsDisposed", 
                    System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance);
                return (bool) propertyInfo.GetValue(this, null);
            }
        }
        #endregion
    }
}
