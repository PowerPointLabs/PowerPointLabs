using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Collections.Specialized;
using System.ComponentModel;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Input;
using PowerPointLabs.ImageSearch.Domain;
using PowerPointLabs.ImageSearch.Handler;
using PowerPointLabs.ImageSearch.SearchEngine;
using PowerPointLabs.ImageSearch.Util;
using PowerPointLabs.Models;
using PowerPointLabs.Utils;
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
        // downloadedImages - to be loaded to SearchList, when button is Download
        // fromFileImages - to be loaded to SearchList, when button is From file
        private List<ImageItem> _downloadedImages;
        private List<ImageItem> _fromFileImages;

        // UI model - list that holds preview item
        public ObservableCollection<ImageItem> PreviewList { get; set; }

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

        // a background presentation that will do the preview processing
        public StylesHandler PreviewPresentation { get; set; }

        // the current image search engine
        public AsyncSearchEngine SearchEngine { get; set; }

        // search engines map
        private readonly Dictionary<string, AsyncSearchEngine> _id2EngineMap 
            = new Dictionary<string, AsyncSearchEngine>(); 

        // indicate whether the window is open/closed or not
        public bool IsOpen { get; set; }

        public StyleOptions StyleOptions { get; set; }

        public SearchOptions SearchOptions { get; set; }

        // indicate whether it's downloading fullsize image, so that debounce.
        // timer - it will download full size image after some time
        // apply - it will download full size image when there's no cache and user clicks APPLY button
        private readonly HashSet<string> _timerDownloadingUriList = new HashSet<string>();
        private readonly HashSet<string> _applyDownloadingUriList = new HashSet<string>();

        private DateTime _latestStyleOptionsUpdateTime = DateTime.Now;
        private DateTime _latestPreviewUpdateTime = DateTime.Now;
        private DateTime _latestPreviewApplyUpdateTime = DateTime.Now;

        private bool _isWindowActivatedWithPreview = true;

        # endregion

        #region Initialization

        public ImageSearchPane()
        {
            InitializeComponent();
            InitSearchTextbox();
            InitMultiplePurposeButtons();
            InitSearchList();
            InitPreviewList();
            IsOpen = true;
            InitStyleOptions();
            InitSearchOptions();
            InitSearchButtons();
            InitConfirmApplyFlyout();
            if (TempPath.InitTempFolder())
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

        private void InitSearchButtons()
        {
            // if no available API Keys, then set default button to Download
            if (SearchOptions.GetSearchEngine() == GoogleEngine.Id()
                && (StringUtil.IsEmpty(SearchOptions.SearchEngineId)
                    || StringUtil.IsEmpty(SearchOptions.ApiKey)))
            {
                SearchButton.SelectedIndex = TextCollection.ImagesLabText.ButtonIndexDownload;
            } 
            else if (SearchOptions.GetSearchEngine() == BingEngine.Id()
                    && StringUtil.IsEmpty(SearchOptions.BingApiKey))
            {
                SearchButton.SelectedIndex = TextCollection.ImagesLabText.ButtonIndexDownload;
            }
        }

        private void InitConfirmApplyFlyout()
        {
            ConfirmApplyPreviewImageFile = new ObservableString { Text = "" };
            ConfirmApplyFlyoutTitle = new ObservableString { Text = "Confirm Apply" };
            ConfirmApplyImage.DataContext = ConfirmApplyPreviewImageFile;
            ConfirmApplyFlyout.DataContext = ConfirmApplyFlyoutTitle;
            ConfirmApplyFlyout.IsOpenChanged += ConfirmApplyFlyout_OnIsOpenChanged;
            OptionsPane2.DataContext = StyleOptions;
        }

        private void InitSearchTextbox()
        {
            SearchTextboxWatermark = new ObservableString {Text = TextCollection.ImagesLabText.TextBoxWatermarkSearch};
            SearchTextBox.DataContext = SearchTextboxWatermark;
        }

        private void InitMultiplePurposeButtons()
        {
            MultiplePurposeButtons = new ObservableCollection<string>(new List<string>
            {
                TextCollection.ImagesLabText.MultiPurposeButtonNameSearch, 
                TextCollection.ImagesLabText.MultiPurposeButtonNameDownload, 
                TextCollection.ImagesLabText.MultiPurposeButtonNameFromFile
            });
            SearchButton.ItemsSource = MultiplePurposeButtons;
            SearchButton.SelectedIndex = TextCollection.ImagesLabText.ButtonIndexSearch;
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
            _downloadedImages = new List<ImageItem>();
            _fromFileImages = new List<ImageItem>();
            SearchList.CollectionChanged += SearchList_OnCollectionChanged;
            SearchListBox.DataContext = this;
        }

        private void InitSearchOptions()
        {
            SearchOptions = SearchOptions.Load(StoragePath.GetPath("ImagesLabSearchOptions"));
            SearchOptions.PropertyChanged += (sender, args) =>
            {
                SearchEngine = _id2EngineMap[SearchOptions.GetSearchEngine()];
            };
            AdvancedPane.DataContext = SearchOptions;
        }

        private void InitStyleOptions()
        {
            StyleOptions = StyleOptions.Load(StoragePath.GetPath("ImagesLabStyleOptions"));
            StyleOptions.PropertyChanged += (sender, args) =>
            {
                _latestStyleOptionsUpdateTime = DateTime.Now;
            };
            OptionsPane.DataContext = StyleOptions;
            StyleOptionsFlyout.IsOpenChanged += StyleOptionsFlyout_OnIsOpenChanged;
        }

        private void InitPreviewPresentation()
        {
            PreviewPresentation = new StylesHandler(StyleOptions);
            PreviewPresentation.Open(withWindow: false, focus: false);
        }

        # endregion

        # region Common UI Events & Interactions
        private void StyleOptionsFlyout_OnIsOpenChanged(object sender, RoutedEventArgs e)
        {
            if (!StyleOptionsFlyout.IsOpen
                && _latestStyleOptionsUpdateTime > _latestPreviewUpdateTime)
            {
                DoPreview();
            }
        }

        private void ConfirmApplyFlyout_OnIsOpenChanged(object sender, RoutedEventArgs e)
        {
            if (!ConfirmApplyFlyout.IsOpen
                && (_latestStyleOptionsUpdateTime > _latestPreviewUpdateTime
                    || _latestPreviewApplyUpdateTime > _latestPreviewUpdateTime))
            {
                DoPreview();
            }
        }

        private void SearchList_OnCollectionChanged(object sender, NotifyCollectionChangedEventArgs e)
        {
            SearchInstructions.Visibility = SearchList.Count == 0
                ? Visibility.Visible
                : Visibility.Hidden;
        }

        private void PreviewList_OnCollectionChanged(object sender, NotifyCollectionChangedEventArgs e)
        {
            if (PreviewList.Count != 0)
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
            if (source != null && source.ImageFile == TempPath.LoadMoreImgPath)
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
                }
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
            ImagesLabGrid.ColumnDefinitions[0].Width = 
                new GridLength(ImagesLabGrid.ColumnDefinitions[0].ActualWidth + e.HorizontalChange);
        }

        // enable & disable insert button
        private void PreivewListBox_OnSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (PreviewListBox.SelectedValue != null)
            {
                PreviewApply.IsEnabled = true;
                ConfirmApplyButton.IsEnabled = true;
                ConfirmApplyPreviewButton.IsEnabled = true;
                UpdateConfirmApplyFlyOutComboBox(PreviewListBox.SelectedItems);
            }
            else
            {
                PreviewApply.IsEnabled = false;
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
            StyleOptions.Save(StoragePath.GetPath("ImagesLabStyleOptions"));
            SearchOptions.Save(StoragePath.GetPath("ImagesLabSearchOptions"));
        }

        private void PreviewApply_OnClick(object sender, RoutedEventArgs e)
        {
            ApplyStyle();
        }

        private void PreviewDisplayToggleSwitch_OnIsCheckedChanged(object sender, EventArgs e)
        {
            Dispatcher.BeginInvoke(new Action(() =>
            {
                var targetColumn = ImagesLabGrid.ColumnDefinitions[0];
                targetColumn.Width = PreviewDisplayToggleSwitch.IsChecked == true 
                    ? new GridLength(620) 
                    : new GridLength(320);
            }));
        }

        private void StyleOptionsButton_OnClick(object sender, RoutedEventArgs e)
        {
            StyleOptionsFlyout.IsOpen = true;
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

            if (ConfirmApplyFlyout.IsOpen)
            {
                switch (e.Key)
                {
                    case Key.Escape:
                        ConfirmApplyFlyout.IsOpen = false;
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
                        if (PreviewApply.IsEnabled)
                        {
                            PreviewApply.RaiseEvent(new RoutedEventArgs(ButtonBase.ClickEvent));
                        }
                        break;
                }
            }
        }

        private void SearchButton_OnSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            SearchList.Clear();
            SearchTextBox.Text = "";
            switch (SearchButton.SelectedIndex)
            {
                case TextCollection.ImagesLabText.ButtonIndexSearch:
                    SearchTextBox.IsEnabled = true;
                    SearchTextboxWatermark.Text = TextCollection.ImagesLabText.TextBoxWatermarkSearch;
                    SearchInstructions.Text = TextCollection.ImagesLabText.InstructionForSearch;
                    FocusSearchTextBox();
                    break;
                case TextCollection.ImagesLabText.ButtonIndexDownload:
                    SearchTextBox.IsEnabled = true;
                    SearchTextboxWatermark.Text = TextCollection.ImagesLabText.TextBoxWatermarkDownload;
                    SearchInstructions.Text = TextCollection.ImagesLabText.InstructionForDownload;
                    CopyContentToObservableList(_downloadedImages, SearchList);
                    break;
                case TextCollection.ImagesLabText.ButtonIndexFromFile:
                    SearchTextBox.IsEnabled = false;
                    SearchTextboxWatermark.Text = TextCollection.ImagesLabText.TextBoxWatermarkFromFile;
                    SearchInstructions.Text = TextCollection.ImagesLabText.InstructionForFromFile;
                    CopyContentToObservableList(_fromFileImages, SearchList);
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

            if (ConfirmApplyFlyout.IsOpen)
            {
                UpdateConfirmApplyPreviewImage();
            }
            else
            {
                DoPreview();
            }
        }

        // intent: clicking 'load more' should not change selection
        private void SearchListBox_OnPreviewMouseDown(object sender, MouseButtonEventArgs e)
        {
            var item = ItemsControl.ContainerFromElement((ItemsControl) sender, (DependencyObject) e.OriginalSource) as ListBoxItem;
            if (item == null || item.Content == null) return;
            var imageItem = item.Content as ImageItem;
            if (imageItem != null && imageItem.ImageFile == TempPath.LoadMoreImgPath)
            {
                DoSearchMore(imageItem);
                e.Handled = true;
            }
        }

        # endregion
    }
}
