using System;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Timers;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Input;
using Microsoft.Office.Core;
using PowerPointLabs.AutoUpdate;
using PowerPointLabs.ImageSearch.Model;
using PowerPointLabs.ImageSearch.VO;
using PowerPointLabs.Models;
using RestSharp;
using RestSharp.Deserializers;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Path = System.IO.Path;

namespace PowerPointLabs.ImageSearch
{
    /// <summary>
    /// Interaction logic for ImageSearchPane.xaml
    /// </summary>
    /// TODO close ppt to close images lab window
    public partial class ImageSearchPane
    {
        public ObservableCollection<ImageItem> SearchList { get; set; }

        public ObservableCollection<ImageItem> PreviewList { get; set; }

        public PowerPointPresentation PreviewPresentation { get; set; }

        private Timer _previewTimer = new Timer { Interval = 1500 };

        private readonly string _loadingImgPath = Path.GetTempPath() + "loading" + DateTime.Now.GetHashCode();

        public ImageSearchPane()
        {
            InitializeComponent();
            // TODO show instructions when lists are empty
            SearchList = new ObservableCollection<ImageItem>();
            PreviewList = new ObservableCollection<ImageItem>();
            SearchListBox.DataContext = this;
            PreviewListBox.DataContext = this;
            // intent: background presentation to do preview processing
            PreviewPresentation = new PowerPointPresentation(Path.GetTempPath(), "imagesLabPreview");
            PreviewPresentation.Open(withWindow: false, focus: false);
            try
            {
                Properties.Resources.Loading.Save(_loadingImgPath);
            }
            catch
            {
                // may fail to save it, cannot override sometimes
            }
            InitPreviewTimer();
        }

        // TODO: 
        // 1. every time sequence different caused by multi-thread (done)
        // 2. when show up the pane, focus on search textbox (done)
        // 3. error handling
        // -- from thread somehow,
        // -- from IO
        // -- from rest (not status code OK)
        // -- from connection
        private void SearchButton_OnClick(object sender, RoutedEventArgs e)
        {
            // TODO: Store this API somewhere...
            var api =
                "https://www.googleapis.com/customsearch/v1?filter=1&cx=017201692871514580973%3Awwdg7q__" +
//                "mb4&imgSize=large&searchType=image&imgType=photo&safe=medium&key=AIzaSyCGcq3O8NN9U7YX-Pj3E7tZde0yaFFeUyY";
//                "mb4&imgSize=large&searchType=image&imgType=photo&safe=medium&key=AIzaSyDQeqy9efF_ASgi2dk3Ortj2QNnz90RdOw";
                "mb4&imgSize=large&searchType=image&imgType=photo&safe=medium&key=AIzaSyDXR8wBYL6al5jXIXTHpEF28CCuvL0fjKk";
            var query = SearchTextBox.Text;
            // TODO: what if query is empty ... may need escape as well

            Dispatcher.BeginInvoke(new Action(() =>
            {
                // intent:
                // clear search list, and show a list of
                // 'Loading...' images
                SearchList.Clear();
                // TODO: number of result needs to be const
                for (int i = 0; i < 30; i++)
                {
                    SearchList.Add(new ImageItem
                    {
                        ImageFile = _loadingImgPath
                    });
                }
                SearchProgressRing.IsActive = true;
            }));

            // TODO the result can be less than 30
            // TODO load more
            SearchImages(api, query, 0);
            SearchImages(api, query, 10);
            SearchImages(api, query, 20, true);
        }

        private void SearchImages(string api, string query, int startIdx, bool isEnd = false)
        {
            var restClient = new RestClient {BaseUrl = new Uri(api + "&start=" + (startIdx + 1) + "&q=" + query)};
            restClient.ExecuteAsync(new RestRequest(Method.GET), response =>
            {
                var deser = new JsonDeserializer();
                var searchResults = deser.Deserialize<SearchResults>(response);
                // TODO: err handling, eg not deser correctly, status code not 200

                for (int i = 0; i < searchResults.Items.Count; i++)
                {
                    var item = SearchList[startIdx + i];
                    var searchResult = searchResults.Items[i];
                    var targetLocation = Path.GetTempPath() + Guid.NewGuid();
                    // intent: 
                    // download thumbnail and show it,
                    // also dump other meta info (e.g. full-size img link)
                    new Downloader()
                        .Get(searchResult.Image.ThumbnailLink, targetLocation)
                        .After(() =>
                        {
                            item.ImageFile = targetLocation;
                            item.FullSizeImageUri = searchResult.Link;
                        })
                        .Start();
                }
                Dispatcher.BeginInvoke(new Action(() =>
                {
                    if (isEnd)
                    {
                        SearchProgressRing.IsActive = false;
                    }
                }));
            });
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
            _previewTimer.Stop();
            DoPreview((ImageItem) SearchListBox.SelectedValue);
            // TODO: start a timer, and if re-select -> reset the timer
            // when timer ticks, try to download full size image to replace
            _previewTimer.Start();
        }

        // intent:
        // when select a thumbnail for some time,
        // try to download its full size version for better preview and can be used for insertion
        private void InitPreviewTimer()
        {
            _previewTimer.Elapsed += (sender, args) =>
            {
                // timer thread
                _previewTimer.Stop();
                Dispatcher.BeginInvoke(new Action(() =>
                {
                    // ui thread
                    if (SearchListBox.SelectedValue != null
                        && (SearchListBox.SelectedValue as ImageItem).FullSizeImageFile == null)
                    {
                        var fullsizeImageFile = Path.GetTempPath() + Guid.NewGuid();
                        var image = (ImageItem) SearchListBox.SelectedValue;
                        new Downloader()
                            .Get(image.FullSizeImageUri, fullsizeImageFile)
                            .After(() =>
                            {
                                // downloader thread
                                Dispatcher.BeginInvoke(new Action(() =>
                                {
                                    // ui thread again
                                    // store back to image, so cache it
                                    image.FullSizeImageFile = fullsizeImageFile;
                                    if (SearchListBox.SelectedValue != null
                                        && (SearchListBox.SelectedValue as ImageItem).ImageFile == image.ImageFile)
                                    {
                                        // intent: aft download, selected value may have been changed
                                        DoPreview(image);
                                    }
                                }));
                            })
                            .Start();
                    }
                }));
            };
        }

        // do preview processing
        private void DoPreview(ImageItem imageItem)
        {
            // ui thread
            Dispatcher.BeginInvoke(new Action(() =>
            {
                PreviewProgressRing.IsActive = true;
                PreviewList.Clear();

                var previewFile = Path.GetTempPath() + "original" + DateTime.Now.GetHashCode();

                // TODO DRY
                var thisSlide = PreviewPresentation.AddSlide(PowerPointCurrentPresentationInfo.CurrentSlide.Layout);
                // TODO has error, when nothing to copy
                PowerPointCurrentPresentationInfo.CurrentSlide.Shapes.Range().Copy();
                thisSlide.Shapes.Paste();
                foreach (PowerPoint.Shape shape in thisSlide.Shapes)
                {
                    if (shape.Name.StartsWith("pptImagesLab"))
                    {
                        shape.Delete();
                    }
                }
                var imageShape = thisSlide.Shapes.AddPicture(imageItem.FullSizeImageFile ?? imageItem.ImageFile, 
                    MsoTriState.msoFalse, MsoTriState.msoTrue, 0,
                    0);
                FitToSlide.AutoFit(imageShape, PreviewPresentation);
                imageShape.ZOrder(MsoZOrderCmd.msoSendToBack);
                thisSlide.GetNativeSlide().Export(previewFile, "JPG");
                // dont affect next time preview
                thisSlide.Delete();

                PreviewList.Add(new ImageItem
                {
                    ImageFile = previewFile,
                    FullSizeImageFile = imageItem.FullSizeImageFile
                });
                // try catch finally?
                PreviewProgressRing.IsActive = false;
            }));
        }


        // intent:
        // allow arrow keys to navigate the search result items in the list
        private void SearchListBox_OnKeyDown(object sender, KeyEventArgs e)
        {
            if (SearchListBox.Items.Count > 0)
            {
                switch (e.Key)
                {
                    case Key.Right:
                    case Key.Down:
                        if (!SearchListBox.Items.MoveCurrentToNext())
                        {
                            SearchListBox.Items.MoveCurrentToLast();
                        }
                        break;

                    case Key.Left:
                    case Key.Up:
                        if (!SearchListBox.Items.MoveCurrentToPrevious())
                        {
                            SearchListBox.Items.MoveCurrentToFirst();
                        }
                        break;

                    default:
                        return;
                }

                e.Handled = true;
                ListBoxItem lbi = (ListBoxItem)SearchListBox.ItemContainerGenerator.ContainerFromItem(SearchListBox.SelectedItem);
                lbi.Focus();
            }
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
                if (PreviewListBox.SelectedValue != null)
                {
                    PreviewInsert.IsEnabled = true;
                }
                else
                {
                    PreviewInsert.IsEnabled = false;
                }
            }));
        }

        // TODO DRY
        private void PreviewListBox_OnKeyDownListBox_OnKeyDown(object sender, KeyEventArgs e)
        {
            if (PreviewListBox.Items.Count > 0)
            {
                switch (e.Key)
                {
                    case Key.Right:
                    case Key.Down:
                        if (!PreviewListBox.Items.MoveCurrentToNext())
                        {
                            PreviewListBox.Items.MoveCurrentToLast();
                        }
                        break;

                    case Key.Left:
                    case Key.Up:
                        if (!PreviewListBox.Items.MoveCurrentToPrevious())
                        {
                            PreviewListBox.Items.MoveCurrentToFirst();
                        }
                        break;

                    default:
                        return;
                }

                e.Handled = true;
                ListBoxItem lbi = (ListBoxItem)PreviewListBox.ItemContainerGenerator.ContainerFromItem(PreviewListBox.SelectedItem);
                lbi.Focus();
            }
        }

        // rmb to close background presentation
        private void ImageSearchPane_OnClosing(object sender, CancelEventArgs e)
        {
            if (PreviewPresentation != null)
            {
                PreviewPresentation.Close();
            }
        }

        // TODO DRY
        private void PreviewInsert_OnClick(object sender, RoutedEventArgs e)
        {
            Dispatcher.BeginInvoke(new Action(() =>
            {
                _previewTimer.Stop();
                PreviewProgressRing.IsActive = true;
            
                // TODO know other style to apply
                // selected value can be null, this works if there's cache for full size image
                if (((ImageItem) SearchListBox.SelectedValue).FullSizeImageFile != null)
                {
                    var thisSlide = PowerPointCurrentPresentationInfo.CurrentSlide;
                    foreach (PowerPoint.Shape shape in thisSlide.Shapes)
                    {
                        if (shape.Name.StartsWith("pptImagesLab"))
                        {
                            shape.Delete();
                        }
                    }
                    var imageShape = thisSlide.Shapes.AddPicture(((ImageItem) PreviewListBox.SelectedValue).FullSizeImageFile, MsoTriState.msoFalse,
                        MsoTriState.msoTrue, 0, 0);
                    imageShape.Name = "pptImagesLab" + DateTime.Now.GetHashCode();
                    FitToSlide.AutoFit(imageShape, PreviewPresentation);
                    imageShape.ZOrder(MsoZOrderCmd.msoSendToBack);
                    PreviewProgressRing.IsActive = false;
                }
                else
                {
                    // download full-size image & apply style's algorithm
                    var imageItem = (ImageItem) SearchListBox.SelectedValue;
                    var fullsizeImageFile = Path.GetTempPath() + Guid.NewGuid();
                    // TODO downloader timeout???
                    new Downloader()
                        .Get(imageItem.FullSizeImageUri, fullsizeImageFile)
                        .After(() =>
                        {
                            Dispatcher.BeginInvoke(new Action(() =>
                            {
                                imageItem.FullSizeImageFile = fullsizeImageFile;
                                if (SearchListBox.SelectedValue != null
                                    && (SearchListBox.SelectedValue as ImageItem).ImageFile == imageItem.ImageFile)
                                {
                                    DoPreview(imageItem);
                                }
                                var thisSlide = PowerPointCurrentPresentationInfo.CurrentSlide;
                                foreach (PowerPoint.Shape shape in thisSlide.Shapes)
                                {
                                    if (shape.Name.StartsWith("pptImagesLab"))
                                    {
                                        shape.Delete();
                                    }
                                }
                                var imageShape = thisSlide.Shapes.AddPicture(fullsizeImageFile, MsoTriState.msoFalse,
                                    MsoTriState.msoTrue, 0, 0);
                                imageShape.Name = "pptImagesLab" + DateTime.Now.GetHashCode();
                                FitToSlide.AutoFit(imageShape, PreviewPresentation);
                                imageShape.ZOrder(MsoZOrderCmd.msoSendToBack);
                                PreviewProgressRing.IsActive = false;
                            }));
                        })
                        .Start();
                }
            }));
        }
    }
}
