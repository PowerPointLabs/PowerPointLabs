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
                if (StringUtil.IsEmpty(SearchOptions.SearchEngineId)
                    || StringUtil.IsEmpty(SearchOptions.ApiKey))
                {
                    ShowErrorMessageBox(TextCollection.ImagesLabText.ErrorNoEngineIdOrApiKey);
                    return;
                }

                SearchButton.IsEnabled = false;
                PrepareToSearch(GoogleEngine.NumOfItemsPerSearch);
                SearchEngine.Search(query);
            }));
        }

        private void DoSearchMore(ImageItem loadMoreItem)
        {
            Dispatcher.Invoke(new Action(() =>
            {
                SearchButton.IsEnabled = false;
                loadMoreItem.ImageFile = TempPath.LoadingImgPath;
                PrepareToSearch(GoogleEngine.NumOfItemsPerRequest - 1, isListClearNeeded: false);
                SearchEngine.SearchMore();
            }));
        }

        private void DoLoadImageFromFile()
        {
            Dispatcher.Invoke(new Action(() =>
            {
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

                try
                {
                    var imgInput = Image.FromFile(openFileDialog.FileName);
                    Graphics.FromImage(imgInput);
                    // so this is an image
                    var fromFileItem = new ImageItem
                    {
                        ImageFile = openFileDialog.FileName,
                        FullSizeImageFile = openFileDialog.FileName,
                        FullSizeImageUri = openFileDialog.FileName,
                        ContextLink = openFileDialog.FileName
                    };
                    SearchList.Add(fromFileItem);
                    _fromFileImages.Add(fromFileItem);
                }
                catch
                {
                    // not an image or image is corrupted
                    ShowErrorMessageBox(TextCollection.ImagesLabText.ErrorImageCorrupted);
                }
            }));
        }

        private void DoDownloadImage()
        {
            Dispatcher.Invoke(new Action(() =>
            {
                var downloadLink = SearchTextBox.Text.Trim();
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
                SearchList.Add(item);
                SearchProgressRing.IsActive = true;

                var thumbnailPath = TempPath.GetPath("thumbnail");
                new Downloader()
                    .Get(downloadLink, thumbnailPath)
                    .After(() =>
                    {
                        HandleDownloadedThumbnail(item, thumbnailPath);
                        HandleDownloadedPicture(item, thumbnailPath);
                    })
                    .OnError(() => { RemoveImageItem(item); })
                    .Start();
            }));
        }

        private void InitSearchEngine()
        {
            SearchEngine = new GoogleEngine(SearchOptions)
                .WhenSucceed(HandleSearchSuccess)
                .WhenCompleted(HandleSearchCompletion)
                .WhenFail(HandleSearchFailure)
                .WhenException(HandleSearchException);
        }

        # endregion

        # region Helper Funcs
        private void HandleSearchSuccess(GoogleSearchResults searchResults, int startIdx)
        {
            // in case null result item
            searchResults.Items = searchResults.Items ?? new List<SearchResult>();
            // in case UI list not prepared
            AddNeededImageItem(startIdx);

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
                    && SearchList.Count + GoogleEngine.NumOfItemsPerRequest - 1 /*loadMore item*/
                        <= GoogleEngine.MaxNumOfItems)
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
            try
            {
                var imgInput = Image.FromFile(thumbnailPath);
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
                RemoveImageItem(item);
            }
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
            while (startIdx + GoogleEngine.NumOfItemsPerRequest - 1 >= SearchList.Count)
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

        private static string GetTooltip(SearchResult searchResult)
        {
            return searchResult.Title + "\n" + searchResult.Image.Width + " x " + searchResult.Image.Height;
        }
        # endregion
    }
}
