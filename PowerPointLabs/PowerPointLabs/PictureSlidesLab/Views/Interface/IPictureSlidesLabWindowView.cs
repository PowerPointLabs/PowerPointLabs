using System;
using System.Threading.Tasks;
using System.Windows.Media;

using MahApps.Metro.Controls.Dialogs;

using PowerPointLabs.PictureSlidesLab.Model;
using PowerPointLabs.PictureSlidesLab.Thread.Interface;

namespace PowerPointLabs.PictureSlidesLab.Views.Interface
{
    public interface IPictureSlidesLabWindowView
    {
        void ShowErrorMessageBox(string content);

        void ShowErrorMessageBox(string content, Exception e);

        Task<MessageDialogResult> ShowInfoMessageBox(string content, MessageDialogStyle dialogStyle);

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
