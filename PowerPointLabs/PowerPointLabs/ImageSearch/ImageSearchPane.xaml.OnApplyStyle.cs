using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using PowerPointLabs.AutoUpdate;
using PowerPointLabs.ImageSearch.Domain;
using PowerPointLabs.ImageSearch.Util;
using PowerPointLabs.Utils;
using ButtonBase = System.Windows.Controls.Primitives.ButtonBase;
using KeyEventArgs = System.Windows.Input.KeyEventArgs;

namespace PowerPointLabs.ImageSearch
{
    public partial class ImageSearchPane
    {
        private void ApplyStyle()
        {
            PreviewTimer.Stop();
            PreviewProgressRing.IsActive = true;

            var source = (ImageItem)SearchListBox.SelectedValue;
            var targetStyle = (ImageItem)PreviewListBox.SelectedValue;
            if (source == null || targetStyle == null) return;

            if (source.FullSizeImageFile != null)
            {
                OpenConfirmApplyFlyout(targetStyle);
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
                    .After(() => { HandleDownloadedFullSizeImage(source, fullsizeImageFile); })
                    .OnError(() =>
                    {
                        Dispatcher.BeginInvoke(new Action(() =>
                        {
                            var currentImageItem = SearchListBox.SelectedValue as ImageItem;
                            if (currentImageItem == null)
                            {
                                PreviewProgressRing.IsActive = false;
                            }
                            else if (currentImageItem.ImageFile == source.ImageFile)
                            {
                                ShowErrorMessageBox(TextCollection.ImagesLabText.ErrorNetworkOrSourceUnavailable);
                            }
                        }));
                    })
                    .Start();
            }
            // already downloading, then update preview image in the map
            else
            {
                var fullsizeImageUri = source.FullSizeImageUri;
                _insertDownloadingUriToPreviewImage[fullsizeImageUri] = targetStyle;
            }
        }

        private void OpenConfirmApplyFlyout(ImageItem targetStyle)
        {
            UpdateConfirmApplyFlyOut(targetStyle);
            ConfirmApplyFlyout.IsOpen = true;
        }

        private void ConfirmApplyPreviewButton_OnClick(object sender, RoutedEventArgs e)
        {
            DoPreview();
        }

        private void ConfirmApplyButton_OnClick(object sender, RoutedEventArgs e)
        {
            var source = SearchListBox.SelectedValue as ImageItem;
            var targetStyle = PreviewListBox.SelectedValue as ImageItem;
            Assumption.Made(source != null && targetStyle != null, "source item or target style item is null");

            PreviewPresentation.ApplyStyle(source, targetStyle.Tooltip);
            ConfirmApplyFlyout.IsOpen = false;
        }

        private void UpdateConfirmApplyFlyOut(ImageItem targetStyle)
        {
            ConfirmApplyFlyoutTitle.Text = "Confirm Apply " + targetStyle.Tooltip;
            switch (targetStyle.Tooltip)
            {
                case TextCollection.ImagesLabText.StyleNameDirectText:
                    TargetStyleComboBox.SelectedIndex = 0;
                    break;
                case TextCollection.ImagesLabText.StyleNameBlur:
                    TargetStyleComboBox.SelectedIndex = 1;
                    break;
                case TextCollection.ImagesLabText.StyleNameTextBox:
                    TargetStyleComboBox.SelectedIndex = 2;
                    break;
                case TextCollection.ImagesLabText.StyleNameBanner:
                    TargetStyleComboBox.SelectedIndex = 3;
                    break;
                // case StyleNameSpecialEffect
                default:
                    TargetStyleComboBox.SelectedIndex = 4;
                    break;
            }
        }

        private void UpdateConfirmApplyPreviewImage()
        {
            if (PreviewListBox.SelectedValue != null)
            {
                var targetImageItem = PreviewListBox.SelectedValue as ImageItem;
                ConfirmApplyPreviewImageFile.Text = targetImageItem.ImageFile;
            }
        }

        private void TargetStyleComboBox_OnSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (PreviewList != null && PreviewList.Count > 0)
            {
                PreviewListBox.SelectedIndex = TargetStyleComboBox.SelectedIndex;
            }
        }

        private void ConfirmApplyFlyout_OnKeyDown(object sender, KeyEventArgs e)
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
    }
}
