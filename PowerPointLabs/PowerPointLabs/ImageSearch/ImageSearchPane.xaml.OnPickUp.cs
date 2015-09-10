using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using MahApps.Metro.Controls.Dialogs;
using PowerPointLabs.AutoUpdate;
using PowerPointLabs.ImageSearch.Domain;
using PowerPointLabs.ImageSearch.Util;
using PowerPointLabs.Utils;
using PowerPointLabs.Utils.Exceptions;

namespace PowerPointLabs.ImageSearch
{
    partial class ImageSearchPane
    {
        private void OpenPickupFlyout()
        {
            UpdateStyleVariationsImages();
            StyleVariationsFlyout.IsOpen = true;
        }

        private void UpdateStyleVariationsImages()
        {
            Dispatcher.BeginInvoke(new Action(() =>
            {
                if (PreviewListBox.SelectedValue == null) return;
            
                var source = SearchListBox.SelectedValue as ImageItem;
                var targetStyleItems = PreviewListBox.SelectedItems;
                var targetStyles = targetStyleItems.Cast<ImageItem>().Select(item => item.Tooltip).ToList();
                Assumption.Made(source != null && targetStyles.Count > 0, "source item or target style item is null/empty");

                try
                {
                    var selectedId = VariationListBox.SelectedIndex >= 0 ? VariationListBox.SelectedIndex : 0;
                    VariationList.Clear();
                    UpdateStyleVariationsImage(StyleOptions, source, targetStyles);
                    UpdateStyleVariationsImage(StyleOptions1, source, targetStyles);
                    UpdateStyleVariationsImage(StyleOptions2, source, targetStyles);
                    UpdateStyleVariationsImage(StyleOptions3, source, targetStyles);
                    UpdateStyleVariationsImage(StyleOptions4, source, targetStyles);
                    UpdateStyleVariationsImage(StyleOptions5, source, targetStyles);
                    VariationListBox.SelectedIndex = selectedId;
                }
                catch
                {
                    // ignore, selected slide may be null
                }
            }));
        }

        private void UpdateStyleVariationsImage(StyleOptions opt, ImageItem source, IList<string> targetStyles)
        {
            PreviewPresentation.SetStyleOptions(opt);
            var previewInfo = PreviewPresentation.PreviewApplyStyle(source, targetStyles);
            VariationList.Add(new ImageItem { ImageFile = previewInfo.PreviewApplyStyleImagePath });
        }

        private void PickUpStyle()
        {
            PreviewTimer.Stop();
            SetProgressingRingStatus(true);

            var source = (ImageItem)SearchListBox.SelectedValue;
            var targetStyle = PreviewListBox.SelectedItems;
            if (source == null || targetStyle == null || targetStyle.Count == 0) return;

            if (source.FullSizeImageFile != null)
            {
                OpenPickupFlyout();
                SetProgressingRingStatus(false);
            }
            else if (!_applyDownloadingUriList.Contains(source.FullSizeImageUri))
            {
                var fullsizeImageUri = source.FullSizeImageUri;
                _applyDownloadingUriList.Add(fullsizeImageUri);

                var fullsizeImageFile = TempPath.GetPath("fullsize");
                new Downloader()
                    .Get(fullsizeImageUri, fullsizeImageFile)
                    .After(() => { HandleDownloadedFullSizeImage(source, fullsizeImageFile); })
                    .OnError(() =>
                    {
                        Dispatcher.BeginInvoke(new Action(() =>
                        {
                            var currentImageItem = SearchListBox.SelectedValue as ImageItem;
                            if (currentImageItem == null)
                            {
                                SetProgressingRingStatus(false);
                            }
                            else if (currentImageItem.ImageFile == source.ImageFile)
                            {
                                ShowErrorMessageBox(TextCollection.ImagesLabText.ErrorNetworkOrSourceUnavailable);
                            }
                        }));
                    })
                    .Start();
            }
        }

        private void VariationListBox_OnSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (VariationListBox.SelectedValue == null)
            {
                StyleApplyButton.IsEnabled = false;
                StyleCustomizeButton.IsEnabled = false;
            }
            else
            {
                StyleApplyButton.IsEnabled = true;
                StyleCustomizeButton.IsEnabled = true;
                switch (VariationListBox.SelectedIndex)
                {
                    case 0:
                        PreviewPresentation.SetStyleOptions(StyleOptions);
                        OptionsPane2.DataContext = StyleOptions;
                        break;
                    case 1:
                        PreviewPresentation.SetStyleOptions(StyleOptions1);
                        OptionsPane2.DataContext = StyleOptions1;
                        break;
                    case 2:
                        PreviewPresentation.SetStyleOptions(StyleOptions2);
                        OptionsPane2.DataContext = StyleOptions2;
                        break;
                    case 3:
                        PreviewPresentation.SetStyleOptions(StyleOptions3);
                        OptionsPane2.DataContext = StyleOptions3;
                        break;
                    case 4:
                        PreviewPresentation.SetStyleOptions(StyleOptions4);
                        OptionsPane2.DataContext = StyleOptions4;
                        break;
                    case 5:
                        PreviewPresentation.SetStyleOptions(StyleOptions5);
                        OptionsPane2.DataContext = StyleOptions5;
                        break;
                }
            }
        }

        private void StyleApplyButton_OnClick(object sender, RoutedEventArgs e)
        {
            if (PreviewListBox.SelectedValue == null) return;

            var source = SearchListBox.SelectedValue as ImageItem;
            var targetStyleItems = PreviewListBox.SelectedItems;
            var targetStyles = targetStyleItems.Cast<ImageItem>().Select(item => item.Tooltip).ToList();
            Assumption.Made(source != null && targetStyles.Count > 0, "source item or target style item is null/empty");

            try
            {
                PreviewPresentation.ApplyStyle(source, targetStyles);
                this.ShowMessageAsync("", TextCollection.ImagesLabText.SuccessfullyAppliedStyle);
            }
            catch (AssumptionFailedException expt)
            {
                ShowErrorMessageBox(TextCollection.ImagesLabText.ErrorNoSelectedSlide);
            }
        }

        private void StyleCustomizeButton_OnClick(object sender, RoutedEventArgs e)
        {
            var source = (ImageItem)SearchListBox.SelectedValue;
            var targetStyles = PreviewListBox.SelectedItems;
            if (source == null || targetStyles == null || targetStyles.Count == 0) return;
            OpenCustomizationFlyout(targetStyles);
        }
    }
}
