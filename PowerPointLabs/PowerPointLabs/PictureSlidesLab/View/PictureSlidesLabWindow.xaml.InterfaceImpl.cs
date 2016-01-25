using System.Windows.Media;
using MahApps.Metro.Controls.Dialogs;
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
            this.ShowMessageAsync("Error", content);
        }

        public void ShowInfoMessageBox(string content)
        {
            this.ShowMessageAsync("Info", content);
        }

        public void ShowSuccessfullyAppliedDialog()
        {
            try
            {
                _gotoSlideDialog.Init("Successfully Applied!");
                _gotoSlideDialog.CustomizeGotoSlideButton("Select",
                    "Select the slide to edit styles.");
                _gotoSlideDialog.FocusOkButton();
                this.ShowMetroDialogAsync(_gotoSlideDialog, MetroDialogOptions);
            }
            catch
            {
                // dialog could be fired multiple times
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
    }
}
