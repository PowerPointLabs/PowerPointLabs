using System.Windows.Media;
using PowerPointLabs.ImagesLab.Thread.Interface;

namespace PowerPointLabs.ImagesLab.View.Interface
{
    public interface IImagesLabWindowView
    {
        void ShowErrorMessageBox(string content);

        void ShowInfoMessageBox(string content);

        void ShowSuccessfullyAppliedDialog();

        IThreadContext GetThreadContext();

        bool IsVariationsFlyoutOpen { get; }

        double GetVariationListBoxScrollOffset();

        void SetVariationListBoxScrollOffset(double offset);

        void SetVariantsColorPanelBackground(Brush color);
    }
}
