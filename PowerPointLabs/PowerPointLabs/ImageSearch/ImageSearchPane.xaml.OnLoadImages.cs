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
                loadMoreItem.ImageFile = StoragePath.LoadingImgPath;
                PrepareToSearch(SearchEngine.NumOfItemsPerRequest() - 1, isListClearNeeded: false);
                SearchEngine.SearchMore();
            }));
        }

        private void DoLoadImageFromFile(string[] filenames = null)
        {
            Dispatcher.Invoke(new Action(() =>
            {
                if (filenames == null)
                {
                    var openFileDialog = new OpenFileDialog
                    {
                        Multiselect = true,
                        Filter = @"Image File|*.png;*.jpg;*.jpeg;*.bmp;*.gif;"
                    };
                    var fileDialogResult = openFileDialog.ShowDialog();
                    if (fileDialogResult != System.Windows.Forms.DialogResult.OK)
                    {
                        return;
                    }
                    filenames = openFileDialog.FileNames;
                }

                try
                {
                    foreach (var filename in filenames)
                    {
                        VerifyIsProperImage(filename);
                        var fromFileItem = new ImageItem
                        {
                            ImageFile = ImageUtil.GetThumbnailFromFullSizeImg(filename),
                            FullSizeImageFile = filename,
                            FullSizeImageUri = filename,
                            ContextLink = filename,
                            Tooltip = ImageUtil.GetWidthAndHeight(filename)
                        };
                        //add it
                        SearchList.Add(fromFileItem);  
                    }

                    if (SearchButton.SelectedIndex != TextCollection.ImagesLabText.ButtonIndexFromFile)
                    {
                        SearchButton.SelectedIndex = TextCollection.ImagesLabText.ButtonIndexFromFile;
                    }
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
                // Error Case 1: If url not valid
                if (!UrlUtil.IsUrlValid(downloadLink))
                {
                    ShowErrorMessageBox(TextCollection.ImagesLabText.ErrorUrlLinkIncorrect);
                    return;
                }

                var item = new ImageItem
                {
                    ImageFile = StoragePath.LoadingImgPath,
                    ContextLink = downloadLink
                };
                UrlUtil.GetMetaInfo(ref downloadLink, item);

                if (SearchButton.SelectedIndex != TextCollection.ImagesLabText.ButtonIndexDownload)
                {
                    SearchButton.SelectedIndex = TextCollection.ImagesLabText.ButtonIndexDownload;
                }
                // add it
                SearchList.Add(item);
                SearchProgressRing.IsActive = true;

                var imagePath = StoragePath.GetPath("img-" 
                    + DateTime.Now.GetHashCode() + "-" 
                    + Guid.NewGuid().ToString().Substring(0, 7));
                new Downloader()
                    .Get(downloadLink, imagePath)
                    .After(() =>
                    {
                        try
                        {
                            // Error Case 2: not a proper image
                            VerifyIsProperImage(imagePath);

                            Dispatcher.Invoke(new Action(() =>
                            {
                                // TODO turn off progress ring after all downloaded
                                SearchProgressRing.IsActive = false;
                            }));
                            HandleDownloadedThumbnail(item, imagePath);
                        }
                        catch
                        {
                            ShowErrorMessageBox(TextCollection.ImagesLabText.ErrorImageCorrupted);
                            Dispatcher.Invoke(new Action(() =>
                            {
                                // TODO turn off progress ring after all downloaded
                                SearchProgressRing.IsActive = false;
                                SearchList.Remove(item);
                            }));
                        }
                    })
                    // Error Case 3: Possibly network timeout
                    .OnError(e => { RemoveImageItem(item, e); })
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
                else if (!isSuccessful && SearchList.Count(source => source.ImageFile == StoragePath.LoadingImgPath)
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

        private void RemoveImageItem(ImageItem item, Exception e)
        {
            Dispatcher.BeginInvoke(new Action(() =>
            {
                SearchProgressRing.IsActive = false;
                SearchList.Remove(item);
            }));
            ShowErrorMessageBox(TextCollection.ImagesLabText.ErrorFailedToLoad + e.Message);
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
                SearchList.Add(new ImageItem { ImageFile = StoragePath.LoadingImgPath });
            }
            SearchProgressRing.IsActive = true;
        }

        private void EnableSearchMore()
        {
            SearchList.Add(new ImageItem
            {
                ImageFile = StoragePath.LoadMoreImgPath
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
                        ImageFile = StoragePath.LoadingImgPath
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
