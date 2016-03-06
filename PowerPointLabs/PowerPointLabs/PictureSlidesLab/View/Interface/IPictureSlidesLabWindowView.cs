using System;
using System.Windows.Media;
using PowerPointLabs.PictureSlidesLab.Model;
using PowerPointLabs.PictureSlidesLab.Thread.Interface;

namespace PowerPointLabs.PictureSlidesLab.View.Interface
{
    public interface IPictureSlidesLabWindowView
    {
        void ShowErrorMessageBox(string content);

        void ShowErrorMessageBox(string content, Exception e);

        void ShowInfoMessageBox(string content);

        void ShowSuccessfullyAppliedDialog();

        IThreadContext GetThreadContext();

        bool IsVariationsFlyoutOpen { get; }

        double GetVariationListBoxScrollOffset();

        void SetVariationListBoxScrollOffset(double offset);

        void SetVariantsColorPanelBackground(Brush color);

        bool IsDisplayDefaultPicture();

        ImageItem CreateDefaultPictureItem();

        void EnterDefaultPictureMode();
    }
}
