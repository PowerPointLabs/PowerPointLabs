using System;
using System.Collections;
using System.Linq;
using System.Windows;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Animation;
using MahApps.Metro.Controls.Dialogs;
using PowerPointLabs.ImageSearch.Domain;
using PowerPointLabs.Utils;
using PowerPointLabs.Utils.Exceptions;
using ButtonBase = System.Windows.Controls.Primitives.ButtonBase;
using KeyEventArgs = System.Windows.Input.KeyEventArgs;

namespace PowerPointLabs.ImageSearch
{
    public partial class ImageSearchPane
    {
        private bool _isCustomizationFlyoutOpen;

        // TODO rename ConfirmApply -> Customization
        private void OpenCustomizationFlyout(IList targetStyles)
        {
            UpdateConfirmApplyPreviewImage();
            _isCustomizationFlyoutOpen = true;

            var toHideTranslate = new TranslateTransform();
            ImagesLabGrid.RenderTransform = toHideTranslate;
            var toHideAnimation = new DoubleAnimation(0, ImagesLabWindow.ActualWidth,
                TimeSpan.FromMilliseconds(600))
            {
                EasingFunction = new SineEase { EasingMode = EasingMode.EaseInOut },
                AccelerationRatio = 0.5
            };
            toHideAnimation.Completed += (sender, args) =>
            {
                Dispatcher.BeginInvoke(new Action(() =>
                {
                    ImagesLabGrid.Visibility = Visibility.Collapsed;
                }));
            };

            var toShowTranslate = new TranslateTransform {X = -ImagesLabWindow.ActualWidth};
            CustomizationFlyout.RenderTransform = toShowTranslate;
            CustomizationFlyout.Visibility = Visibility.Visible;
            var toShowAnimation = new DoubleAnimation(-ImagesLabWindow.ActualWidth, 0,
                TimeSpan.FromMilliseconds(600))
            {
                EasingFunction = new SineEase { EasingMode = EasingMode.EaseInOut },
                AccelerationRatio = 0.5
            };

            toHideTranslate.BeginAnimation(TranslateTransform.XProperty, toHideAnimation);
            toShowTranslate.BeginAnimation(TranslateTransform.XProperty, toShowAnimation);
        }

        private void CustomizationFlyoutBackButton_OnClick(object sender, RoutedEventArgs e)
        {
            CloseCustomizationFlyout();
        }

        private void CloseCustomizationFlyout()
        {
            _isCustomizationFlyoutOpen = false;

            var toShowTranslate = new TranslateTransform { X = ImagesLabWindow.ActualWidth };
            ImagesLabGrid.RenderTransform = toShowTranslate;
            ImagesLabGrid.Visibility = Visibility.Visible;
            var toShowAnimation = new DoubleAnimation(ImagesLabWindow.ActualWidth, 0,
                TimeSpan.FromMilliseconds(600))
            {
                EasingFunction = new SineEase { EasingMode = EasingMode.EaseInOut },
                AccelerationRatio = 0.5
            };

            var toHideTranslate = new TranslateTransform { X = 0 };
            CustomizationFlyout.RenderTransform = toHideTranslate;
            var toHideAnimation = new DoubleAnimation(0, -ImagesLabWindow.ActualWidth,
                TimeSpan.FromMilliseconds(600))
            {
                EasingFunction = new SineEase { EasingMode = EasingMode.EaseInOut },
                AccelerationRatio = 0.5
            };
            toHideAnimation.Completed += (sender, args) =>
            {
                Dispatcher.BeginInvoke(new Action(() =>
                {
                    CustomizationFlyout.Visibility = Visibility.Collapsed;
                }));
            };

            toShowTranslate.BeginAnimation(TranslateTransform.XProperty, toShowAnimation);
            toHideTranslate.BeginAnimation(TranslateTransform.XProperty, toHideAnimation);
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
                PreviewPresentation.ApplyStyle(source);
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
            catch (AssumptionFailedException)
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
                    var previewInfo = PreviewPresentation.PreviewApplyStyle(source, isActualSize:true);

                    ConfirmApplyPreviewImageFile.Text = previewInfo.PreviewApplyStyleImagePath;
                    _latestPreviewApplyUpdateTime = DateTime.Now;
                }
                catch
                {
                    // ignore, selected slide may be null
                }
            }));
        }

        private void CustomizationFlyout_OnKeyDown(object sender, KeyEventArgs e)
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
    }
}
