using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using PowerPointLabs.AutoUpdate;
using PowerPointLabs.ImageSearch.Domain;
using PowerPointLabs.ImageSearch.SearchEngine;
using PowerPointLabs.ImageSearch.SearchEngine.VO;
using PowerPointLabs.ImageSearch.Util;
using PowerPointLabs.Utils;
using RestSharp;

namespace PowerPointLabs.ImageSearch
{
    public partial class ImageSearchPane
    {
        # region Internal APIs

        private void DoSearch()
        {
            Dispatcher.Invoke(new Action(() =>
            {
                var query = SearchTextBox.Text;
                if (StringUtil.IsEmpty(query))
                {
                    return;
                }
                if (SearchOptions.GetSearchEngine() == TextCollection.ImagesLabText.SearchEngineGoogle 
                    && (StringUtil.IsEmpty(SearchOptions.SearchEngineId)
                        || StringUtil.IsEmpty(SearchOptions.ApiKey)))
                {
                    ShowErrorMessageBox(TextCollection.ImagesLabText.ErrorNoEngineIdOrApiKey);
                    return;
                }
                if (SearchOptions.GetSearchEngine() == TextCollection.ImagesLabText.SearchEngineBing
                    && StringUtil.IsEmpty(SearchOptions.BingApiKey))
                {
                    ShowErrorMessageBox(TextCollection.ImagesLabText.ErrorNoEngineIdOrApiKey);
                    return;
                }

                SearchButton.IsEnabled = false;
                PrepareToSearch(SearchEngine.NumOfItemsPerSearch());
                SearchEngine.Search(query);
            }));
        }

        private void DoSearchMore(ImageItem loadMoreItem)
        {
            Dispatcher.Invoke(new Action(() =>
            {
                SearchButton.IsEnabled = false;
                loadMoreItem.ImageFile = TempPath.LoadingImgPath;
                PrepareToSearch(SearchEngine.NumOfItemsPerRequest() - 1, isListClearNeeded: false);
                SearchEngine.SearchMore();
            }));
        }

        private void DoLoadImageFromFile(string filename = null)
        {
            Dispatcher.Invoke(new Action(() =>
            {
                if (filename == null)
                {
                    var openFileDialog = new OpenFileDialog
                    {
                        Multiselect = false,
                        Filter = @"Image File|*.png;*.jpg;*.jpeg;*.bmp;*.gif;"
                    };
                    var fileDialogResult = openFileDialog.ShowDialog();
                    if (fileDialogResult != System.Windows.Forms.DialogResult.OK)
                    {
                        return;
                    }
                    filename = openFileDialog.FileName;
                }

                try
                {
                    VerifyIsProperImage(filename);
                    var fromFileItem = new ImageItem
                    {
                        ImageFile = filename,
                        FullSizeImageFile = filename,
                        FullSizeImageUri = filename,
                        ContextLink = filename
                    };

                    if (SearchButton.SelectedIndex != TextCollection.ImagesLabText.ButtonIndexFromFile)
                    {
                        SearchButton.SelectedIndex = TextCollection.ImagesLabText.ButtonIndexFromFile;
                    }
                    //add and select it
                    SearchList.Add(fromFileItem);
                    SearchListBox.SelectedIndex = SearchListBox.Items.Count - 1;
                    _fromFileImages.Add(fromFileItem);
                }
                catch
                {
                    // not an image or image is corrupted
                    ShowErrorMessageBox(TextCollection.ImagesLabText.ErrorImageCorrupted);
                }
            }));
        }

        private static void VerifyIsProperImage(string filename)
        {
            using (Image.FromFile(filename))
            {
                // so this is a proper image
            }
        }

        private void DoDownloadImage(string downloadLink = null)
        {
            Dispatcher.Invoke(new Action(() =>
            {
                downloadLink = downloadLink ?? SearchTextBox.Text.Trim();
                if (StringUtil.IsEmpty(downloadLink))
                {
                    return;
                }
                if (!UrlUtil.IsUrlValid(downloadLink))
                {
                    ShowErrorMessageBox(TextCollection.ImagesLabText.ErrorUrlLinkIncorrect);
                    return;
                }

                var item = new ImageItem
                {
                    ImageFile = TempPath.LoadingImgPath,
                    ContextLink = downloadLink
                };
                UrlUtil.GetMetaInfo(ref downloadLink, item);

                var selectedIds = GetSelectedIndices(PreviewListBox.SelectedItems);
                if (SearchButton.SelectedIndex != TextCollection.ImagesLabText.ButtonIndexDownload)
                {
                    SearchButton.SelectedIndex = TextCollection.ImagesLabText.ButtonIndexDownload;
                }
                // add and select it
                SearchList.Add(item);
                SearchListBox.SelectedIndex = SearchListBox.Items.Count - 1;
                SearchProgressRing.IsActive = true;

                var thumbnailPath = TempPath.GetPath("thumbnail");
                new Downloader()
                    .Get(downloadLink, thumbnailPath)
                    .After(() =>
                    {
                        HandleDownloadedPicture(item, thumbnailPath);
                        HandleDownloadedThumbnail(item, thumbnailPath, selectedIds: selectedIds);
                    })
                    .OnError(() => { RemoveImageItem(item); })
                    .Start();
            }));
        }

        private void InitSearchEngine()
        {
            var googleEngine = new GoogleEngine(SearchOptions)
                .WhenSucceed(HandleSearchSuccess)
                .WhenCompleted(HandleSearchCompletion)
                .WhenFail(HandleSearchFailure)
                .WhenException(HandleSearchException);
            var bingEngine = new BingEngine(SearchOptions)
                .WhenSucceed(HandleSearchSuccess)
                .WhenCompleted(HandleSearchCompletion)
                .WhenFail(HandleSearchFailure)
                .WhenException(HandleSearchException);

            _id2EngineMap.Add(GoogleEngine.Id(), googleEngine);
            _id2EngineMap.Add(BingEngine.Id(), bingEngine);
            SearchEngine = _id2EngineMap[SearchOptions.GetSearchEngine()];
        }

        # endregion

        # region Helper Funcs
        private void HandleSearchSuccess(object results, int startIdx)
        {
            dynamic searchResults = results;
            // in case null result item
            if (results is GoogleSearchResults)
            {
                searchResults.Items = searchResults.Items ?? new List<GoogleSearchResult>();
            }
            else if (results is BingSearchResults)
            {
                searchResults.D = searchResults.D ?? 
                    new BingSearchResultsWrapper { Results = new List<BingSearchResult>() };
                searchResults.D.Results = searchResults.D.Results ?? new List<BingSearchResult>();
            }
            else return;

            // in case UI list not prepared
            AddNeededImageItem(startIdx);

            for (var i = 0; i < SearchEngine.NumOfItemsPerRequest(); i++)
            {
                var item = SearchList[startIdx + i];
                if (i >= VOUtil.GetCount(searchResults))
                {
                    item.IsToDelete = true;
                    continue;
                }

                object searchResult = VOUtil.GetItem(searchResults, i);
                var thumbnailPath = TempPath.GetPath("thumbnail");

                new Downloader()
                    .Get(VOUtil.GetThumbnailLink(searchResult), thumbnailPath)
                    .After(()=> { HandleDownloadedThumbnail(item, thumbnailPath, searchResult); })
                    .Start();
            }
        }

        private void HandleSearchCompletion(bool isSuccessful)
        {
            Dispatcher.BeginInvoke(new Action(() =>
            {
                SearchProgressRing.IsActive = false;
                var isThereMoreSearchResults = !RemoveElementInSearchList(item => item.IsToDelete);
                SearchButton.IsEnabled = true;

                if (isSuccessful
                    && isThereMoreSearchResults
                    && SearchList.Count + SearchEngine.NumOfItemsPerRequest() - 1 /*loadMore item*/
                        <= SearchEngine.MaxNumOfItems())
                {
                    EnableSearchMore();
                }
                // all failed then clear list
                else if (!isSuccessful && SearchList.Count(source => source.ImageFile == TempPath.LoadingImgPath)
                        == SearchList.Count)
                {
                    SearchList.Clear();
                }
            }));
        }

        private void HandleSearchFailure(IRestResponse response)
        {
            ShowErrorMessageBox(TextCollection.ImagesLabText.ErrorNetworkOrApiQuotaUnavailable);
        }

        private void HandleSearchException(Exception e)
        {
            ShowErrorMessageBox(e.Message);
        }

        private void RemoveImageItem(ImageItem item)
        {
            Dispatcher.BeginInvoke(new Action(() =>
            {
                SearchProgressRing.IsActive = false;
                SearchList.Remove(item);
            }));
            ShowErrorMessageBox(TextCollection.ImagesLabText.ErrorImageCorrupted);
        }

        private void HandleDownloadedPicture(ImageItem item, string thumbnailPath)
        {
            VerifyIsProperImage(thumbnailPath);
            Dispatcher.Invoke(new Action(() =>
            {
                SearchProgressRing.IsActive = false;
                _downloadedImages.Add(item);
            }));
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

        private void EnableSearchMore()
        {
            SearchList.Add(new ImageItem
            {
                ImageFile = TempPath.LoadMoreImgPath
            });
        }

        private void AddNeededImageItem(int startIdx)
        {
            while (startIdx + SearchEngine.NumOfItemsPerRequest() - 1 >= SearchList.Count)
            {
                Dispatcher.Invoke(new Action(() =>
                {
                    SearchList.Add(new ImageItem
                    {
                        ImageFile = TempPath.LoadingImgPath
                    });
                }));
            }
        }

        private bool RemoveElementInSearchList(Func<ImageItem, bool> cond)
        {
            var isAnyElementRemoved = false;
            for (var i = 0; i < SearchList.Count; )
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

        private static string GetTooltip(object searchResult)
        {
            return VOUtil.GetTitle(searchResult) + "\n" + VOUtil.GetWidth(searchResult) + " x " + VOUtil.GetHeight(searchResult);
        }
        # endregion
    }
}
