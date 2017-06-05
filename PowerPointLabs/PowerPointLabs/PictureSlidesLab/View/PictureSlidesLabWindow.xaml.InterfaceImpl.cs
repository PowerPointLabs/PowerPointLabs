using System;
using System.Threading.Tasks;
using System.Windows.Media;
using MahApps.Metro.Controls.Dialogs;
using PowerPointLabs.ActionFramework.Common.Log;
using PowerPointLabs.PictureSlidesLab.Model;
using PowerPointLabs.PictureSlidesLab.Thread;
using PowerPointLabs.PictureSlidesLab.Thread.Interface;
using PowerPointLabs.PictureSlidesLab.Util;

namespace PowerPointLabs.PictureSlidesLab.View
{
    public partial class PictureSlidesLabWindow
    {
        ///////////////////////////////////////////////////////////////
        // Implemented interface methods
        ///////////////////////////////////////////////////////////////

        public void ShowErrorMessageBox(string content)
        {
            try
            {
                this.ShowMessageAsync("Error", content);
            }
            catch (Exception e)
            {
                Logger.LogException(e, "ShowErrorMessageBox");
            }
        }

        /// <summary>
        /// Show msgBox for exception
        /// </summary>
        /// <param name="content"></param>
        /// <param name="e"></param>
        public void ShowErrorMessageBox(string content, Exception e)
        {
            if (e == null)
            {
                ShowErrorMessageBox(content);
                return;
            }

            try
            {
                ShowErrorTextDialog(content + TextCollection.UserFeedBack + TextCollection.Email + "\r\n\r\n"
                                               + e.Message + " " + e.GetType() + "\r\n"
                                               + e.StackTrace);
            }
            catch (Exception expt)
            {
                Logger.LogException(e, "ShowErrorMessageBox (parameter)");
                Logger.LogException(expt, "ShowErrorMessageBox");
            }
        }

        public Task<MessageDialogResult> ShowInfoMessageBox(string content, MessageDialogStyle dialogStyle = MessageDialogStyle.Affirmative)
        {
            return this.ShowMessageAsync("Info", content, dialogStyle);
        }

        public void ShowSuccessfullyAppliedDialog()
        {
            try
            {
                if (_gotoSlideDialog.IsOpen)
                {
                    return;
                }

                _gotoSlideDialog
                    .Init("Successfully Applied!")
                    .CustomizeGotoSlideButton("Select", "Select the slide to edit styles.")
                    .FocusOkButton()
                    .OpenDialog();
                this.ShowMetroDialogAsync(_gotoSlideDialog, MetroDialogOptions);
            }
            catch (Exception e)
            {
                Logger.LogException(e, "ShowSuccessfullyAppliedDialog");
            }
        }

        public IThreadContext GetThreadContext()
        {
            return new ThreadContext(Dispatcher);
        }

        public double GetVariationListBoxScrollOffset()
        {
            var scrollOffset = 0d;
            var scrollViewer = ListBoxUtil.FindScrollViewer(StylesVariationListBox);
            if (scrollViewer != null) { scrollOffset = scrollViewer.VerticalOffset; }
            return scrollOffset;
        }

        public void SetVariationListBoxScrollOffset(double offset)
        {
            var scrollViewer = ListBoxUtil.FindScrollViewer(StylesVariationListBox);
            if (scrollViewer != null) { scrollViewer.ScrollToVerticalOffset(offset); }
        }

        public void SetVariantsColorPanelBackground(Brush color)
        {
            VariantsColorPanel.Background = color;
        }

        public ImageItem CreateDefaultPictureItem()
        {
            return new ImageItem
            {
                ImageFile = StoragePath.NoPicturePlaceholderImgPath,
                Tooltip = "Please select a picture."
            };
        }

        public bool IsDisplayDefaultPicture()
        {
            return _isDisplayDefaultPicture;
        }

        public void EnterDefaultPictureMode()
        {
            _isDisplayDefaultPicture = true;
            _isEnableUpdatePreview = false;
            ImageSelectionListBox.SelectedIndex = -1;
            _isEnableUpdatePreview = true;
        }
    }
}
