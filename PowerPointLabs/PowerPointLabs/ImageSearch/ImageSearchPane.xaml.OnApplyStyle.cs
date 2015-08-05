using System;
using System.Collections;
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
            var targetStyle = PreviewListBox.SelectedItems;
            if (source == null || targetStyle == null || targetStyle.Count == 0) return;

            if (source.FullSizeImageFile != null)
            {
                OpenConfirmApplyFlyout(targetStyle);
                PreviewProgressRing.IsActive = false;
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
        }

        private void OpenConfirmApplyFlyout(IList targetStyles)
        {
            UpdateConfirmApplyPreviewImage();
            UpdateConfirmApplyFlyOutComboBox(targetStyles);
            ConfirmApplyFlyout.IsOpen = true;
        }

        private void ConfirmApplyPreviewButton_OnClick(object sender, RoutedEventArgs e)
        {
            UpdateConfirmApplyPreviewImage();
        }

        private void ConfirmApplyButton_OnClick(object sender, RoutedEventArgs e)
        {
            if (PreviewListBox.SelectedValue == null) return;

            var source = SearchListBox.SelectedValue as ImageItem;
            var targetStyleItems = PreviewListBox.SelectedItems;
            var targetStyles = targetStyleItems.Cast<ImageItem>().Select(item => item.Tooltip).ToList();
            Assumption.Made(source != null && targetStyles.Count > 0, "source item or target style item is null/empty");

            try
            {
                PreviewPresentation.ApplyStyle(source, targetStyles);
                this.ShowMessageAsync("", TextCollection.ImagesLabText.SuccessfullyAppliedStyle)
                    .ContinueWith(task =>
                    {
                        Dispatcher.BeginInvoke(new Action(() =>
                        {
                            if (_latestStyleOptionsUpdateTime > _latestPreviewApplyUpdateTime)
                            {
                                UpdateConfirmApplyPreviewImage();
                            }
                            ConfirmApplyButton.Focus();
                            Keyboard.Focus(ConfirmApplyButton);
                        }));
                    });
            }
            catch (AssumptionFailedException expt)
            {
                ShowErrorMessageBox(TextCollection.ImagesLabText.ErrorNoSelectedSlide);
            }
        }

        private void UpdateConfirmApplyPreviewImage()
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
                    var previewInfo = PreviewPresentation.PreviewApplyStyle(source, targetStyles);

                    ConfirmApplyPreviewImageFile.Text = previewInfo.PreviewApplyStyleImagePath;
                    _latestPreviewApplyUpdateTime = DateTime.Now;
                }
                catch
                {
                    // ignore, selected slide may be null
                }
            }));
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

        # region Sync PreviewListBox styles selection & CheckBox styles selection

        // Sync PreviewListBox styles selection to CheckBox styles selection
        private void UpdateConfirmApplyFlyOutComboBox(IList targetStyles)
        {
            TickCheckBox(
                GetCheckBoxFromComboBoxItem(TargetStyleComboBox.Items[TextCollection.ImagesLabText.StyleIndexDirectText]),
                HasStyle(targetStyles, TextCollection.ImagesLabText.StyleNameDirectText));
            TickCheckBox(
                GetCheckBoxFromComboBoxItem(TargetStyleComboBox.Items[TextCollection.ImagesLabText.StyleIndexBlur]),
                HasStyle(targetStyles, TextCollection.ImagesLabText.StyleNameBlur));
            TickCheckBox(
                GetCheckBoxFromComboBoxItem(TargetStyleComboBox.Items[TextCollection.ImagesLabText.StyleIndexTextBox]),
                HasStyle(targetStyles, TextCollection.ImagesLabText.StyleNameTextBox));
            TickCheckBox(
                GetCheckBoxFromComboBoxItem(TargetStyleComboBox.Items[TextCollection.ImagesLabText.StyleIndexBanner]),
                HasStyle(targetStyles, TextCollection.ImagesLabText.StyleNameBanner));
            TickCheckBox(
                GetCheckBoxFromComboBoxItem(TargetStyleComboBox.Items[TextCollection.ImagesLabText.StyleIndexSpecialEffect]),
                HasStyle(targetStyles, TextCollection.ImagesLabText.StyleNameSpecialEffect));
        }

        // Sync CheckBox styles selection to PreviewListBox styles selection (when checked)
        private void TargetStyleCheckBox_OnChecked(object sender, RoutedEventArgs e)
        {
            var targetStyleCheckBox = sender as CheckBox;
            if (targetStyleCheckBox == null) return;

            SyncCheckBoxSelectionToPreviewListBox(targetStyleCheckBox);
        }

        // Sync CheckBox styles selection to PreviewListBox styles selection (when unchecked)
        private void TargetStyleCheckBox_OnUnchecked(object sender, RoutedEventArgs e)
        {
            var targetStyleCheckBox = sender as CheckBox;
            if (targetStyleCheckBox == null) return;

            SyncCheckBoxSelectionToPreviewListBox(targetStyleCheckBox, isToAddSelection: false);
        }

        // the rest is helper func

        private bool HasStyle(IList targetStyles, string style)
        {
            return targetStyles.Cast<ImageItem>().Any(targetStyle => targetStyle.Tooltip == style);
        }

        private CheckBox GetCheckBoxFromComboBoxItem(Object item)
        {
            return (item as ComboBoxItem) != null ? ((ComboBoxItem)item).Content as CheckBox : null;
        }

        private void TickCheckBox(CheckBox cb, bool isChecked)
        {
            if (cb == null) return;
            var originalTooltip = cb.ToolTip as string;
            // avoid triggering Checked/Unchecked event
            cb.ToolTip = "Updating...";
            cb.IsChecked = isChecked;
            cb.ToolTip = originalTooltip;
        }

        private void SyncCheckBoxSelectionToPreviewListBox(CheckBox targetStyleCheckBox, bool isToAddSelection = true)
        {
            switch (targetStyleCheckBox.ToolTip as string)
            {
                case "Style 1":
                    SelectPreviewListBox(TextCollection.ImagesLabText.StyleIndexDirectText, isToAddSelection);
                    break;
                case "Style 2":
                    SelectPreviewListBox(TextCollection.ImagesLabText.StyleIndexBlur, isToAddSelection);
                    break;
                case "Style 3":
                    SelectPreviewListBox(TextCollection.ImagesLabText.StyleIndexTextBox, isToAddSelection);
                    break;
                case "Style 4":
                    SelectPreviewListBox(TextCollection.ImagesLabText.StyleIndexBanner, isToAddSelection);
                    break;
                case "Style 5":
                    SelectPreviewListBox(TextCollection.ImagesLabText.StyleIndexSpecialEffect, isToAddSelection);
                    break;
            }
        }

        private void SelectPreviewListBox(int index, bool isToAdd)
        {
            if (index >= PreviewListBox.Items.Count) return;
            if (isToAdd)
            {
                PreviewListBox.SelectedItems.Add(PreviewListBox.Items[index]);
            }
            else
            {
                PreviewListBox.SelectedItems.Remove(PreviewListBox.Items[index]);
            }
        }

        #endregion
    }
}
