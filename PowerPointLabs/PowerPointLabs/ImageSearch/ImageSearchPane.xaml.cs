using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Collections.Specialized;
using System.ComponentModel;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Animation;
using PowerPointLabs.ImageSearch.Domain;
using PowerPointLabs.ImageSearch.Handler;
using PowerPointLabs.ImageSearch.SearchEngine;
using PowerPointLabs.ImageSearch.Util;
using PowerPointLabs.Models;
using PowerPointLabs.WPF.Observable;
using ButtonBase = System.Windows.Controls.Primitives.ButtonBase;
using KeyEventArgs = System.Windows.Input.KeyEventArgs;
using ListBox = System.Windows.Controls.ListBox;
using Timer = System.Timers.Timer;

namespace PowerPointLabs.ImageSearch
{
    /// <summary>
    /// Interaction logic for ImageSearchPane.xaml
    /// </summary>
    /// TODO to do unit test for WPF UI,
    /// MVVM pattern must be applied
    public partial class ImageSearchPane
    {
        # region Props & States

        // UI model - list that holds search result item
        public ObservableCollection<ImageItem> SearchList { get; set; }

        // caches for multiple-purpose buttons
        // downloadedImages - to be loaded to SearchList, when button is Download/From File
        private List<ImageItem> _downloadedImages;

        // UI model - list that holds preview item
        public ObservableCollection<ImageItem> PreviewList { get; set; }

        // UI model - list that holds style variations item
        public ObservableCollection<ImageItem> VariationList { get; set; }

        // UI model - list that holds multiple purpose buttons
        public ObservableCollection<string> MultiplePurposeButtons { get; set; }

        // UI model - search textbox watermark
        public ObservableString SearchTextboxWatermark { get; set; }

        // UI model - preview image in Confirm Apply flyout
        public ObservableString ConfirmApplyPreviewImageFile { get; set; }

        // UI model
        public ObservableString ConfirmApplyFlyoutTitle { get; set; }

        // UI model for drag and drop instructions
        public ObservableString DragAndDropInstructionText { get; set; }

        // a timer used to download full-size image at background
        public Timer PreviewTimer { get; set; }

        // time to trigger the timer event
        private const int TimerInterval = 2000;
        private const string ImagesLabImagesList = "ImagesLabImagesList";

        // a background presentation that will do the preview processing
        public StylesHandler PreviewPresentation { get; set; }

        // the current image search engine
        public AsyncSearchEngine SearchEngine { get; set; }

        // search engines map
        private readonly Dictionary<string, AsyncSearchEngine> _id2EngineMap 
            = new Dictionary<string, AsyncSearchEngine>(); 

        // indicate whether the window is open/closed or not
        public bool IsOpen { get; set; }

        public SearchOptions SearchOptions { get; set; }

        // indicate whether it's downloading fullsize image, so that debounce.
        // timer - it will download full size image after some time
        // apply - it will download full size image when there's no cache and user clicks APPLY button
        private readonly HashSet<string> _timerDownloadingUriList = new HashSet<string>();
        private readonly HashSet<string> _applyDownloadingUriList = new HashSet<string>();
        private readonly HashSet<string> _customizeDownloadingUriList = new HashSet<string>();

        private DateTime _latestStyleOptionsUpdateTime = DateTime.Now;
        private DateTime _latestPreviewApplyUpdateTime = DateTime.Now;

        private DateTime _latestPreviewUpdateTime = DateTime.Now;
        private DateTime _latestImageChangedTime = DateTime.Now;

        private int _rightClickedSearchListBoxItemIndex = -1;

        private bool _isWindowActivatedWithPreview = true;
        private bool _isStylePreviewRegionInit;
        private bool _isStylePreviewRegionAnimationCompleted;

        # endregion

        #region Initialization

        public ImageSearchPane()
        {
            InitializeComponent();
            InitSearchTextbox();
            InitMultiplePurposeButtons();
            InitSearchList();
            InitPreviewList();
            InitVariationList();
            IsOpen = true;
            InitSearchOptions();
            InitConfirmApplyFlyout();
            if (TempPath.InitTempFolder() && StoragePath.InitPersistentFolder())
            {
                InitSearchEngine();
                InitPreviewPresentation();
                InitPreviewTimer();
                InitDragAndDrop();
            }
            else
            {
                ShowErrorMessageBox(TextCollection.ImagesLabText.ErrorFailToInitTempFolder);
            }
        }

        private void InitVariationList()
        {
            VariationList = new ObservableCollection<ImageItem>();
            VariationList.CollectionChanged += VariationList_OnCollectionChanged;
            VariationListBox.DataContext = this;
        }

        private void InitConfirmApplyFlyout()
        {
            ConfirmApplyPreviewImageFile = new ObservableString { Text = "" };
            ConfirmApplyFlyoutTitle = new ObservableString { Text = "Confirm Apply" };
            ConfirmApplyImage.DataContext = ConfirmApplyPreviewImageFile;
            CustomizationFlyout.DataContext = ConfirmApplyFlyoutTitle;
        }

        private void InitSearchTextbox()
        {
            SearchTextboxWatermark = new ObservableString {Text = TextCollection.ImagesLabText.TextBoxWatermarkFromFile};
            SearchTextBox.DataContext = SearchTextboxWatermark;
        }

        private void InitMultiplePurposeButtons()
        {
            MultiplePurposeButtons = new ObservableCollection<string>(new List<string>
            {
                TextCollection.ImagesLabText.MultiPurposeButtonNameDownload, 
                TextCollection.ImagesLabText.MultiPurposeButtonNameFromFile
            });
            SearchButton.ItemsSource = MultiplePurposeButtons;
            SearchButton.SelectedIndex = TextCollection.ImagesLabText.ButtonIndexFromFile;
        }

        private void InitPreviewList()
        {
            PreviewList = new ObservableCollection<ImageItem>();
            PreviewList.CollectionChanged += PreviewList_OnCollectionChanged;
            PreviewListBox.DataContext = this;
        }

        private void InitSearchList()
        {
            SearchList = new ObservableCollection<ImageItem>();
            _downloadedImages = StoragePath.Load(ImagesLabImagesList);
            SearchList.CollectionChanged += SearchList_OnCollectionChanged;
            CopyContentToObservableList(_downloadedImages, SearchList);
            Dispatcher.BeginInvoke(new Action(() =>
            {
                SearchListBox.DataContext = this;
            }));
        }

        private void InitSearchOptions()
        {
            SearchOptions = new SearchOptions();
            SearchOptions.PropertyChanged += (sender, args) =>
            {
                SearchEngine = _id2EngineMap[SearchOptions.GetSearchEngine()];
            };
            AdvancedPane.DataContext = SearchOptions;
        }

        private void InitPreviewPresentation()
        {
            PreviewPresentation = new StylesHandler();
            PreviewPresentation.Open(withWindow: false, focus: false);
        }

        # endregion

        # region Common UI Events & Interactions

        private void SearchList_OnCollectionChanged(object sender, NotifyCollectionChangedEventArgs e)
        {
            SearchInstructions.Visibility = SearchList.Count == 0
                ? Visibility.Visible
                : Visibility.Hidden;

            // show StylesPreviewRegion aft there'r some images in the SearchList region
            if (SearchList.Count > 0 && !_isStylePreviewRegionInit)
            {
                // only one entry
                _isStylePreviewRegionInit = true;
                var isPreviewInstructionsVisible = PreviewInstructions.Visibility == Visibility.Visible;
                PreviewInstructions.Visibility = Visibility.Hidden;
                PreviewInstructions.Opacity = 0;
                var isPreviewInstructionsWhenNoSelectedSlideVisible =
                    PreviewInstructionsWhenNoSelectedSlide.Visibility == Visibility.Visible;
                PreviewInstructionsWhenNoSelectedSlide.Visibility = Visibility.Hidden;
                PreviewInstructionsWhenNoSelectedSlide.Opacity = 0;
                
                var previewRegionShowAnimation = new DoubleAnimation(0, 560d, TimeSpan.FromMilliseconds(600))
                {
                    EasingFunction = new SineEase { EasingMode = EasingMode.EaseInOut },
                    AccelerationRatio = 0.5
                };

                StylesPreviewGrid.Visibility = Visibility.Visible;
                previewRegionShowAnimation.Completed += (o, args) =>
                {
                    Dispatcher.BeginInvoke(new Action(() =>
                    {
                        var previewInstructionsShowAnimation = 
                            new DoubleAnimation(0, 1, TimeSpan.FromMilliseconds(250))
                        {
                            EasingFunction = new SineEase { EasingMode = EasingMode.EaseInOut },
                            AccelerationRatio = 0.5
                        };

                        if (isPreviewInstructionsVisible)
                        {
                            PreviewInstructions.Visibility = Visibility.Visible;
                            PreviewInstructions.BeginAnimation(OpacityProperty, previewInstructionsShowAnimation);
                        }
                        else if (isPreviewInstructionsWhenNoSelectedSlideVisible)
                        {
                            PreviewInstructionsWhenNoSelectedSlide.Visibility = Visibility.Visible;
                            PreviewInstructionsWhenNoSelectedSlide.BeginAnimation(OpacityProperty,
                                previewInstructionsShowAnimation);
                        }
                        _isStylePreviewRegionAnimationCompleted = true;
                    }));
                };
                StylesPreviewGrid.BeginAnimation(WidthProperty, previewRegionShowAnimation);
            }
        }

        private void VariationList_OnCollectionChanged(object sender, NotifyCollectionChangedEventArgs e)
        {
            if (VariationList.Count != 0)
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

        private void PreviewList_OnCollectionChanged(object sender, NotifyCollectionChangedEventArgs e)
        {
            if (!_isStylePreviewRegionAnimationCompleted) return;

            PreviewInstructions.BeginAnimation(OpacityProperty, null);
            PreviewInstructions.Opacity = 100;
            PreviewInstructionsWhenNoSelectedSlide.BeginAnimation(OpacityProperty, null);
            PreviewInstructionsWhenNoSelectedSlide.Opacity = 100;

            if (PreviewList.Count != 0 || SearchInstructions.Visibility == Visibility.Visible)
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

        private void SearchButton_OnClick(object sender, RoutedEventArgs e)
        {
            switch (SearchButton.SelectedIndex)
            {
                case TextCollection.ImagesLabText.ButtonIndexSearch:
                    CloseVariationsFlyout();
                    DoSearch();
                    break;
                case TextCollection.ImagesLabText.ButtonIndexDownload:
                    DoDownloadImage();
                    break;
                case TextCollection.ImagesLabText.ButtonIndexFromFile:
                    DoLoadImageFromFile();
                    break;
            }
        }

        // intent:
        // press Enter in the textbox to start searching
        private void SearchTextBox_OnKeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                SearchButton_OnClick(sender, e);
                SearchTextBox.SelectAll();
            }
        }

        // intent:
        // do previewing, when search result item is (not) selected
        private void SearchListBox_OnSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            var source = (ImageItem) SearchListBox.SelectedValue;
            if (source != null && source.ImageFile == StoragePath.LoadMoreImgPath)
            {
                DoSearchMore(source);
            }
            else
            {
                // when selection changed, no need to insert 
                // but dont clear timerDownloadingUriList, since timer may still downloading
                // full size image at the background.
                if (source != null)
                {
                    _applyDownloadingUriList.Remove(source.FullSizeImageUri);
                    _customizeDownloadingUriList.Remove(source.FullSizeImageUri);
                }
                else
                {
                    CloseFlyouts();
                }
                _latestImageChangedTime = DateTime.Now;
                DoPreview();
            }
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

        private void SearchListBox_OnKeyDown(object sender, KeyEventArgs e)
        {
            switch (e.Key)
            {
                case Key.Delete:
                case Key.Back:
                    DeleteImageShape();
                    return;
            }
            ListBox_OnKeyDown(sender, e);
        }

        // intent: focus on search textbox when
        // pane is open
        public void FocusSearchTextBox()
        {
            SearchTextBox.Focus();
            SearchTextBox.SelectAll();
        }

        // intent: drag splitter to change grid width
        private void Splitter_OnDragDelta(object sender, DragDeltaEventArgs e)
        {
            // need remove animation before set its width
            StylesPreviewGrid.BeginAnimation(WidthProperty, null);
            StylesPreviewGrid.Width = StylesPreviewGrid.ActualWidth + e.HorizontalChange;
        }

        // enable & disable insert button
        private void PreivewListBox_OnSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (PreviewListBox.SelectedValue != null)
            {
                StylesPickUpButton.IsEnabled = true;
                ConfirmApplyButton.IsEnabled = true;
                ConfirmApplyPreviewButton.IsEnabled = true;
            }
            else
            {
                StylesPickUpButton.IsEnabled = false;
                ConfirmApplyButton.IsEnabled = false;
                ConfirmApplyPreviewButton.IsEnabled = false;
            }
        }

        // rmb to close background presentation
        private void ImageSearchPane_OnClosing(object sender, CancelEventArgs e)
        {
            IsOpen = false;
            if (PreviewPresentation != null)
            {
                PreviewPresentation.Close();
            }
            StoragePath.Save(ImagesLabImagesList, _downloadedImages);
        }

        private void StylesPickUpButton_OnClick(object sender, RoutedEventArgs e)
        {
            PickUpStyle();
        }

        private void AdvancedButton_OnClick(object sender, RoutedEventArgs e)
        {
            SearchOptionsFlyout.IsOpen = true;
        }

        private void PreviewListBox_OnKeyUp(object sender, KeyEventArgs e)
        {
            var listbox = sender as ListBox;
            if (listbox == null || listbox.Items.Count <= 0)
            {
                return;
            }

            if (_isCustomizationFlyoutOpen)
            {
                switch (e.Key)
                {
                    case Key.Escape:
                        CloseCustomizationFlyout();
                        break;
                    case Key.Enter:
                        ConfirmApplyButton.RaiseEvent(new RoutedEventArgs(ButtonBase.ClickEvent));
                        break;
                }
            }
            else
            {
                switch (e.Key)
                {
                    case Key.Enter:
                        if (StylesPickUpButton.IsEnabled)
                        {
                            StylesPickUpButton.RaiseEvent(new RoutedEventArgs(ButtonBase.ClickEvent));
                        }
                        break;
                }
            }
        }

        private void SearchButton_OnSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            CloseVariationsFlyout();
            SearchTextBox.Text = "";
            switch (SearchButton.SelectedIndex)
            {
                case TextCollection.ImagesLabText.ButtonIndexSearch:
                    SearchListBoxContextMenu.Visibility = Visibility.Collapsed;
                    SearchTextBox.IsEnabled = true;
                    SearchTextboxWatermark.Text = TextCollection.ImagesLabText.TextBoxWatermarkSearch;
                    SearchInstructions.Text = TextCollection.ImagesLabText.InstructionForSearch;
                    FocusSearchTextBox();
                    break;
                case TextCollection.ImagesLabText.ButtonIndexDownload:
                    SearchListBoxContextMenu.Visibility = Visibility.Visible;
                    SearchTextBox.Visibility = Visibility.Visible;
                    SearchTextboxWatermark.Text = TextCollection.ImagesLabText.TextBoxWatermarkDownload;
                    SearchInstructions.Text = TextCollection.ImagesLabText.InstructionForDownload;
                    break;
                case TextCollection.ImagesLabText.ButtonIndexFromFile:
                    SearchListBoxContextMenu.Visibility = Visibility.Visible;
                    SearchTextBox.Visibility = Visibility.Collapsed;
                    SearchTextboxWatermark.Text = TextCollection.ImagesLabText.TextBoxWatermarkFromFile;
                    SearchInstructions.Text = TextCollection.ImagesLabText.InstructionForFromFile;
                    break;
            }
        }

        private void CopyContentToObservableList(IEnumerable<ImageItem> content, ObservableCollection<ImageItem> target)
        {
            foreach (var image in content)
            {
                target.Add(new ImageItem
                {
                    ImageFile = image.ImageFile,
                    FullSizeImageFile = image.FullSizeImageFile,
                    FullSizeImageUri = image.FullSizeImageUri,
                    ContextLink = image.ContextLink
                });
            }
        }

        private void ImageSearchPane_OnActivated(object sender, EventArgs e)
        {
            if (!_isWindowActivatedWithPreview) return;

            Dispatcher.BeginInvoke(new Action(() =>
            {
                if (_isCustomizationFlyoutOpen)
                {
                    UpdateConfirmApplyPreviewImage();
                }
                else if (_isVariationsFlyoutOpen)
                {
                    UpdateStyleVariationsImages();
                }
                else
                {
                    DoPreview();
                }
            }));
        }

        private void SearchListBox_OnPreviewMouseDown(object sender, MouseButtonEventArgs e)
        {
            var item = ItemsControl.ContainerFromElement((ItemsControl) sender, (DependencyObject) e.OriginalSource) as ListBoxItem;
            if (item == null || item.Content == null) return;
            var imageItem = item.Content as ImageItem;
            // intent: clicking 'load more' should not change selection
            if (imageItem != null && imageItem.ImageFile == StoragePath.LoadMoreImgPath)
            {
                DoSearchMore(imageItem);
                e.Handled = true;
            }
            // intent: right click SearchListBox item should not change selection
            else if (e.RightButton == MouseButtonState.Pressed)
            {
                _rightClickedSearchListBoxItemIndex = -1;
                for (var i = 0; i < SearchListBox.Items.Count; i++)
                {
                    var lbi = SearchListBox.ItemContainerGenerator.ContainerFromIndex(i) as ListBoxItem;
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
            _downloadedImages.Clear();
            SearchList.Clear();
        }

        private void DeleteImageShape()
        {
            if (SearchButton.SelectedIndex == TextCollection.ImagesLabText.ButtonIndexSearch) return;
            
            if (_rightClickedSearchListBoxItemIndex < 0 
                || _rightClickedSearchListBoxItemIndex > SearchListBox.Items.Count)
                return;

            var selectedImage = (ImageItem) SearchListBox.Items.GetItemAt(_rightClickedSearchListBoxItemIndex);
            if (selectedImage == null) return;

            if (selectedImage.ImageFile != StoragePath.LoadingImgPath)
            {
                switch (SearchButton.SelectedIndex)
                {
                    case TextCollection.ImagesLabText.ButtonIndexDownload:
                    case TextCollection.ImagesLabText.ButtonIndexFromFile:
                        if (SearchListBox.SelectedIndex < _downloadedImages.Count
                            && SearchListBox.SelectedIndex >= 0)
                        {
                            _downloadedImages.RemoveAt(_rightClickedSearchListBoxItemIndex);
                        }
                        break;
                }
            }
            SearchList.RemoveAt(_rightClickedSearchListBoxItemIndex);
        }

        # endregion
    }
}
