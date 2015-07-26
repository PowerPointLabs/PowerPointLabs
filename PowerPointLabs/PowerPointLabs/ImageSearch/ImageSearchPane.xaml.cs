using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Drawing;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Forms;
using System.Windows.Input;
using MahApps.Metro.Controls.Dialogs;
using PowerPointLabs.AutoUpdate;
using PowerPointLabs.ImageSearch.Domain;
using PowerPointLabs.ImageSearch.Handler;
using PowerPointLabs.ImageSearch.SearchEngine;
using PowerPointLabs.ImageSearch.SearchEngine.VO;
using PowerPointLabs.ImageSearch.Util;
using PowerPointLabs.Models;
using PowerPointLabs.WPF.Observable;
using ButtonBase = System.Windows.Controls.Primitives.ButtonBase;
using Clipboard = System.Windows.Clipboard;
using Image = System.Drawing.Image;
using KeyEventArgs = System.Windows.Input.KeyEventArgs;
using ListBox = System.Windows.Controls.ListBox;
using Timer = System.Timers.Timer;

namespace PowerPointLabs.ImageSearch
{
    // TODO more partial class to clean up codes
    /// <summary>
    /// Interaction logic for ImageSearchPane.xaml
    /// </summary>
    public partial class ImageSearchPane
    {
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

        // a timer used to download full-size image at background
        public Timer PreviewTimer { get; set; }

        // time to trigger the timer event
        private const int TimerInterval = 2000;

        // a background presentation that will do the preview processing
        public StylesHandler PreviewPresentation { get; set; }

        // the image search engine
        public GoogleEngine SearchEngine { get; set; }

        // indicate whether the window is open/closed or not
        public bool IsOpen { get; set; }

        public StyleOptions StyleOptions { get; set; }

        public SearchOptions SearchOptions { get; set; }

        // indicate whether it's downloading fullsize image, so that debounce.
        // timer - it will download full size image after some time
        // insert - it will download full size image when there's no cache and user clicks insert button
        private readonly HashSet<string> _timerDownloadingUriList = new HashSet<string>();
        private readonly HashSet<string> _insertDownloadingUriList = new HashSet<string>();
        private readonly Dictionary<string, ImageItem> _insertDownloadingUriToPreviewImage = new Dictionary<string, ImageItem>();

        // TODO put to text collection
        private const string ErrorNetworkOrApiQuotaUnavailable =
            "Failed to search images. Please check your network, or the daily API quota is ran out.";

        private const string ErrorNetworkOrSourceUnavailable =
            "Failed to insert style. Please check your network, or the image source is unavailable.";

        private const string ErrorNoEngineIdOrApiKey =
            "Please fill in Search Engine Id and API Key by clicking ADVANCED.. button.";

        #region Initialization
        public ImageSearchPane()
        {
            InitializeComponent();
            InitSearchTextbox();
            InitMultiplePurposeButtons();
            InitSearchList();
            InitPreviewList();
            SearchListBox.DataContext = this;
            PreviewListBox.DataContext = this;
            IsOpen = true;
            InitStyleOptions();
            InitSearchOptions();

            var isTempFolderReady = TempPath.InitTempFolder();
            if (isTempFolderReady)
            {
                InitSearchEngine();
                InitPreviewPresentation();
                InitPreviewTimer();
            }
        }

        private void InitSearchTextbox()
        {
            SearchTextboxWatermark = new ObservableString {Text = "Type here to search for images"};
            SearchTextBox.DataContext = SearchTextboxWatermark;
        }

        private void InitMultiplePurposeButtons()
        {
            MultiplePurposeButtons = new ObservableCollection<string>(new List<string> { "Search", "Download", "From File" });
            SearchButton.ItemsSource = MultiplePurposeButtons;
            SearchButton.SelectedIndex = 0;
        }

        private void InitPreviewList()
        {
            PreviewList = new ObservableCollection<ImageItem>();
            PreviewList.CollectionChanged += (sender, args) =>
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
            };
        }

        private void InitSearchList()
        {
            SearchList = new ObservableCollection<ImageItem>();
            _downloadedImages = new List<ImageItem>();
            _fromFileImages = new List<ImageItem>();
            SearchList.CollectionChanged += (sender, args) =>
            {
                SearchInstructions.Visibility = SearchList.Count == 0 
                    ? Visibility.Visible 
                    : Visibility.Hidden;
            };
        }

        private void InitSearchOptions()
        {
            SearchOptions = SearchOptions.Load(StoragePath.GetPath("ImagesLabSearchOptions"));
            AdvancedPane.DataContext = SearchOptions;
        }

        private void InitStyleOptions()
        {
            StyleOptions = StyleOptions.Load(StoragePath.GetPath("ImagesLabStyleOptions"));
            OptionsPane.DataContext = StyleOptions;
            StyleOptionsFlyout.IsOpenChanged += (sender, args) =>
            {
                if (!StyleOptionsFlyout.IsOpen)
                {
                    DoPreview();
                }
            };
        }

        private void InitSearchEngine()
        {
            SearchEngine = new GoogleEngine(SearchOptions)
                .WhenSucceed(WhenSearchSucceed())
                .WhenCompleted(WhenSearchCompleted())
                .WhenFail(response => {
                    ShowErrorMessageBox(ErrorNetworkOrApiQuotaUnavailable);
                    Dispatcher.BeginInvoke(new Action(() =>
                    {
                        // if all failed
                        if (SearchList.Count(source => source.ImageFile == TempPath.LoadingImgPath) 
                            == SearchList.Count)
                        {
                            SearchList.Clear();
                        }
                    }));
                });
        }

        private void ShowErrorMessageBox(string content)
        {
            Dispatcher.BeginInvoke(new Action(() =>
            {
                this.ShowMessageAsync("Error", content);
            }));
        }

        private GoogleEngine.WhenCompletedEventDelegate WhenSearchCompleted()
        {
            return isSuccess =>
            {
                Dispatcher.BeginInvoke(new Action(() =>
                {
                    SearchProgressRing.IsActive = false;
                    var isThereMoreSearchResults = !RemoveElementInSearchList(item => item.IsToDelete);
                    SearchButton.IsEnabled = true;
                    if (isSuccess 
                        && isThereMoreSearchResults 
                        && SearchList.Count + GoogleEngine.NumOfItemsPerRequest - 1/*loadMore item*/ 
                        <= GoogleEngine.MaxNumOfItems)
                    {
                        SearchList.Add(new ImageItem
                        {
                            ImageFile = TempPath.LoadMoreImgPath
                        });
                    }
                }));
            };
        }

        // TODO util
        private bool RemoveElementInSearchList(Func<ImageItem, bool> cond)
        {
            var isAnyElementRemoved = false;
            for (var i = 0; i < SearchList.Count;)
            {
                if (cond(SearchList[i]))
                {
                    SearchList.RemoveAt(i);
                    isAnyElementRemoved = true;
                }
                else
                {
                    i++;
                }
            }
            return isAnyElementRemoved;
        }

        private GoogleEngine.WhenSucceedEventDelegate WhenSearchSucceed()
        {
            return (searchResults, startIdx) =>
            {
                searchResults.Items = searchResults.Items ?? new List<SearchResult>();
                for (var i = 0; i < GoogleEngine.NumOfItemsPerRequest; i++)
                {
                    var item = SearchList[startIdx + i];
                    if (i >= searchResults.Items.Count)
                    {
                        item.IsToDelete = true;
                        continue;
                    }

                    var searchResult = searchResults.Items[i];
                    var thumbnailPath = TempPath.GetPath("thumbnail");

                    new Downloader()
                        .Get(searchResult.Image.ThumbnailLink, thumbnailPath)
                        .After(AfterDownloadThumbnail(item, thumbnailPath, searchResult))
                        .Start();
                }
            };
        }

        private Downloader.AfterDownloadEventDelegate AfterDownloadThumbnail(
            ImageItem item, string thumbnailPath, SearchResult searchResult)
        {
            return () =>
            {
                if (item == null) return;
                item.ImageFile = thumbnailPath;
                if (searchResult != null)
                {
                    item.FullSizeImageUri = searchResult.Link;
                    item.Tooltip = GetTooltip(searchResult);
                    item.ContextLink = searchResult.Image.ContextLink;
                }
                else // use case when thumbnail is already full-size
                {
                    item.FullSizeImageFile = item.ImageFile;
                }

                Dispatcher.BeginInvoke(new Action(() =>
                {
                    var selectedImageItem = SearchListBox.SelectedValue as ImageItem;
                    if (selectedImageItem != null && item.ImageFile == selectedImageItem.ImageFile)
                    {
                        DoPreview();
                    }
                }));
            };
        }

        private static string GetTooltip(SearchResult searchResult)
        {
            return searchResult.Title + "\n" + searchResult.Image.Width + " x " + searchResult.Image.Height;
        }

        private void InitPreviewPresentation()
        {
            PreviewPresentation = new StylesHandler(StyleOptions);
            PreviewPresentation.Open(withWindow: false, focus: false);
        }

        // intent:
        // when select a thumbnail for some time (defined by TimerInterval),
        // try to download its full size version for better preview and can be used for insertion
        private void InitPreviewTimer()
        {
            PreviewTimer = new Timer { Interval = TimerInterval };
            PreviewTimer.Elapsed += (sender, args) =>
            {
                // in timer thread
                PreviewTimer.Stop();
                Dispatcher.BeginInvoke(new Action(() =>
                {
                    // UI thread starts
                    var source = SearchListBox.SelectedValue as ImageItem;
                    // if already have cached full-size image, ignore
                    if (source == null || source.FullSizeImageFile != null)
                    {
                        // do nothing
                    }
                    // if not downloading the full size image yet, download it
                    else if (!_timerDownloadingUriList.Contains(source.FullSizeImageUri))
                    {
                        _timerDownloadingUriList.Add(source.FullSizeImageUri);
                        // preview progress ring will be off, after preview processing is done
                        PreviewProgressRing.IsActive = true;

                        var fullsizeImageFile = TempPath.GetPath("fullsize");
                        new Downloader()
                            .Get(source.FullSizeImageUri, fullsizeImageFile)
                            .After(AfterDownloadFullSizeImage(source, fullsizeImageFile))
                            .OnError(WhenFailDownloadFullSizeImage())
                            .Start();
                    }
                    // it's downloading
                    else
                    {
                        // preview progress ring will be off, after preview processing is done
                        PreviewProgressRing.IsActive = true;
                    }
                }));
            };
        }

        private Downloader.ErrorEventDelegate WhenFailDownloadFullSizeImage()
        {
            return () =>
            {
                // in downloader thread
                Dispatcher.BeginInvoke(new Action(() =>
                {
                    PreviewProgressRing.IsActive = false;
                }));
            };
        }

        private Downloader.AfterDownloadEventDelegate
            AfterDownloadFullSizeImage(ImageItem source, string fullsizeImageFile)
        {
            // timer's downloading will come here at the end,
            // or both timer + insert's downloading will come here
            return () =>
            {
                // in downloader thread
                Dispatcher.BeginInvoke(new Action(() =>
                {
                    if (source == null) return;

                    // UI thread again
                    // store back to image, so cache it
                    source.FullSizeImageFile = fullsizeImageFile;
                    var fullsizeImageUri = source.FullSizeImageUri;

                    // intent: during download, selected item may have been changed to another one
                    // if selected one got changed,
                    // 1. no need to preview it
                    // 2. no need to insert it to current slide
                    var currentImageItem = SearchListBox.SelectedValue as ImageItem;
                    if (currentImageItem == null)
                    {
                        PreviewProgressRing.IsActive = false;
                    }
                    else if (currentImageItem.ImageFile == source.ImageFile)
                    {
                        // if selected one remains
                        // and it is to insert the full size image,
                        ImageItem targetStyle;
                        if (_insertDownloadingUriList.Contains(fullsizeImageUri)
                            && _insertDownloadingUriToPreviewImage
                                .TryGetValue(fullsizeImageUri, out targetStyle))
                        {
                            // insert + do preview
                            PreviewPresentation.ApplyStyle(source, targetStyle.Tooltip);
                            DoPreview(source);
                        } 
                        // or it is to preview only (from timer)
                        else if (_timerDownloadingUriList.Contains(fullsizeImageUri))
                        {
                            DoPreview(source);
                        }
                    }

                    RemoveDebounceCheck(fullsizeImageUri);
                }));
            };
        }

        private void RemoveDebounceCheck(string fullsizeImageUri)
        {
            if (_insertDownloadingUriList.Remove(fullsizeImageUri))
            {
                _insertDownloadingUriToPreviewImage.Remove(fullsizeImageUri);
            }
            _timerDownloadingUriList.Remove(fullsizeImageUri);
        }

        # endregion

        private void SearchButton_OnClick(object sender, RoutedEventArgs e)
        {
            switch (SearchButton.SelectedIndex)
            {
                case 0:// search
                    StartSearching();
                    break;
                case 1:// download link
                    var downloadLink = SearchTextBox.Text.Trim();
                    if (downloadLink.Length == 0)
                    {
                        return;
                    }

                    // TODO parse downloadLink.. 
                    // maybe a link from google image?
                    // maybe without http://
                    var item = new ImageItem
                    {
                        ImageFile = TempPath.LoadingImgPath,
                        ContextLink = downloadLink
                    };
                    SearchList.Add(item);
                    SearchProgressRing.IsActive = true;

                    var thumbnailPath = TempPath.GetPath("thumbnail");
                    new Downloader()
                        .Get(downloadLink, thumbnailPath)
                        .After(() =>
                        {
                            AfterDownloadThumbnail(item, thumbnailPath, null)();
                            try
                            {
                                Image imgInput = Image.FromFile(thumbnailPath);
                                Graphics.FromImage(imgInput);
                                // so this is an image
                                Dispatcher.Invoke(new Action(() =>
                                {
                                    SearchProgressRing.IsActive = false;
                                    _downloadedImages.Add(item);
                                }));
                            }
                            catch
                            {
                                // not an image
                                Dispatcher.BeginInvoke(new Action(() =>
                                {
                                    SearchProgressRing.IsActive = false;
                                    SearchList.Remove(item);
                                }));
                            }
                        })
                        .OnError(() =>
                        {
                            Dispatcher.BeginInvoke(new Action(() =>
                            {
                                SearchProgressRing.IsActive = false;
                                SearchList.Remove(item);
                            }));
                            
                        })
                        .Start();
                    break;
                case 2:// image from file
                    var openFileDialog = new OpenFileDialog
                    {
                        DefaultExt = "*.png",
                        Multiselect = false,
                        Filter = @"Image File|*.png;*.jpg;*.jpeg;"
                    };
                    var fileDialogResult = openFileDialog.ShowDialog();
                    if (fileDialogResult != System.Windows.Forms.DialogResult.OK)
                    {
                        return;
                    }

                    var fromFileItem = new ImageItem
                    {
                        ImageFile = openFileDialog.FileName,
                        FullSizeImageFile = openFileDialog.FileName,
                        FullSizeImageUri = openFileDialog.FileName,
                        ContextLink = openFileDialog.FileName
                    };
                    SearchList.Add(fromFileItem);
                    _fromFileImages.Add(fromFileItem);
                    break;
            }
        }

        private void StartSearching()
        {
            var query = SearchTextBox.Text;
            if (query.Trim().Length == 0)
            {
                return;
            }
            else if (SearchOptions.SearchEngineId.Trim().Length == 0
                     || SearchOptions.ApiKey.Trim().Length == 0)
            {
                ShowErrorMessageBox(ErrorNoEngineIdOrApiKey);
                return;
            }

            SearchButton.IsEnabled = false;
            PrepareToSearch(GoogleEngine.NumOfItemsPerSearch);
            SearchEngine.Search(query);
        }

        private void PrepareToSearch(int expectedNumOfImages, bool isListClearNeeded = true)
        {
            // clear search list, and show a list of
            // 'Loading...' images
            if (isListClearNeeded)
            {
                SearchList.Clear();
            }
            for (var i = 0; i < expectedNumOfImages; i++)
            {
                SearchList.Add(new ImageItem { ImageFile = TempPath.LoadingImgPath });
            }
            SearchProgressRing.IsActive = true;
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
            if (source == null || source.ImageFile == TempPath.LoadingImgPath)
            {
                PreviewList.Clear();
                PreviewProgressRing.IsActive = false;
            } 
            else if (source.ImageFile == TempPath.LoadMoreImgPath)
            {
                PreviewList.Clear();
                PreviewProgressRing.IsActive = false;

                SearchButton.IsEnabled = false;
                source.ImageFile = TempPath.LoadingImgPath;
                PrepareToSearch(GoogleEngine.NumOfItemsPerRequest - 1, isListClearNeeded: false);
                SearchEngine.SearchMore();
            }
            else
            {
                // when selection changed, no need to insert style
                if (_insertDownloadingUriList.Remove(source.FullSizeImageUri))
                {
                    _insertDownloadingUriToPreviewImage.Remove(source.FullSizeImageUri);
                }
                // but dont clear timerDownloadingUriList, since timer may still downloading
                // full size image at the background.

                PreviewTimer.Stop();

                DoPreview(source);

                // when timer ticks, try to download full size image to replace
                PreviewTimer.Start();
            }
        }

        // do preview processing
        private void DoPreview(ImageItem source)
        {
            // ui thread
            Dispatcher.BeginInvoke(new Action(() =>
            {
                var previousTextCopy = Clipboard.GetText();
                var selectedId = PreviewListBox.SelectedIndex;
                PreviewList.Clear();

                if (PowerPointCurrentPresentationInfo.CurrentSlide != null)
                {
                    var previewInfo = PreviewPresentation.PreviewStyles(source);
                    Add(PreviewList, previewInfo.DirectTextStyleImagePath, TextCollection.ImagesLabText.StyleNameDirectText);
                    Add(PreviewList, previewInfo.BlurStyleImagePath, TextCollection.ImagesLabText.StyleNameBlur);
                    Add(PreviewList, previewInfo.TextboxStyleImagePath, TextCollection.ImagesLabText.StyleNameTextBox);
                    Add(PreviewList, previewInfo.GrayScaleStyleImagePath, TextCollection.ImagesLabText.StyleNameGrayscale);

                    PreviewListBox.SelectedIndex = selectedId;
                }
                if (previousTextCopy.Length > 0)
                {
                    Clipboard.SetText(previousTextCopy);
                }
                PreviewProgressRing.IsActive = false;
            }));
        }

        // TODO util
        private void Add(ICollection<ImageItem> list, string imagePath, string tooltip)
        {
            list.Add(new ImageItem
            {
                ImageFile = imagePath,
                Tooltip = tooltip
            });
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
            var item = (ListBoxItem) listbox.ItemContainerGenerator.ContainerFromItem(listbox.SelectedItem);
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
            ImagesLabGrid.ColumnDefinitions[0].Width = new GridLength(ImagesLabGrid.ColumnDefinitions[0].ActualWidth + e.HorizontalChange);
        }

        // enable & disable insert button
        private void PreivewListBox_OnSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            Dispatcher.BeginInvoke(new Action(() =>
            {
                PreviewInsert.IsEnabled = PreviewListBox.SelectedValue != null;
            }));
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

        private void PreviewInsert_OnClick(object sender, RoutedEventArgs e)
        {
            PreviewTimer.Stop();
            PreviewProgressRing.IsActive = true;

            var source = (ImageItem) SearchListBox.SelectedValue;
            var targetStyle = (ImageItem) PreviewListBox.SelectedValue;
            if (source == null || targetStyle == null) return;

            if (source.FullSizeImageFile != null)
            {
                PreviewPresentation.ApplyStyle(source, targetStyle.Tooltip);
                PreviewProgressRing.IsActive = false;
            }
            else if (!_insertDownloadingUriList.Contains(source.FullSizeImageUri))
            {
                var fullsizeImageUri = source.FullSizeImageUri;
                _insertDownloadingUriList.Add(fullsizeImageUri);
                _insertDownloadingUriToPreviewImage[fullsizeImageUri] = targetStyle;

                var fullsizeImageFile = TempPath.GetPath("fullsize");
                new Downloader()
                    .Get(fullsizeImageUri, fullsizeImageFile)
                    .After(AfterDownloadFullSizeImage(source, fullsizeImageFile))
                    .OnError(() => { ShowErrorMessageBox(ErrorNetworkOrSourceUnavailable); })
                    .Start();
            }
            // already downloading, then update preview image in the map
            else
            {
                var fullsizeImageUri = source.FullSizeImageUri;
                _insertDownloadingUriToPreviewImage[fullsizeImageUri] = targetStyle;
            }
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

        private void DoPreview()
        {
            var image = (ImageItem) SearchListBox.SelectedValue;
            if (image == null || image.ImageFile == TempPath.LoadingImgPath)
            {
                PreviewList.Clear();
                PreviewProgressRing.IsActive = false;
            }
            else
            {
                PreviewTimer.Stop();
                DoPreview(image);
                PreviewTimer.Start();
            }
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

            switch (e.Key)
            {
                case Key.Enter:
                    if (PreviewInsert.IsEnabled)
                    {
                        PreviewInsert.RaiseEvent(new RoutedEventArgs(ButtonBase.ClickEvent));
                    }
                    break;
            }
        }

        private void SearchButton_OnSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            SearchList.Clear();
            SearchTextBox.Text = "";
            switch (SearchButton.SelectedIndex)
            {
                // TODO remove those magic
                case 0: // search
                    SearchTextBox.IsEnabled = true;
                    SearchTextboxWatermark.Text = "Type here to search for images";
                    SearchInstructions.Text = "No result. Type the keywords in the textbox above to search for images.";
                    break;
                case 1: // download
                    SearchTextBox.IsEnabled = true;
                    SearchTextboxWatermark.Text = "Paste the image link here";
                    SearchInstructions.Text = "No image. Paste the image link in the textbox above to download images.";
                    CopyContentToObservableList(_downloadedImages, SearchList);
                    break;
                case 2: // from file
                    SearchTextBox.IsEnabled = false;
                    SearchTextboxWatermark.Text = "";
                    SearchInstructions.Text = "No image. Click the 'From File' button above to load local images.";
                    CopyContentToObservableList(_fromFileImages, SearchList);
                    break;
            }
        }

        // TODO util
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
            DoPreview();
        }
    }
}
