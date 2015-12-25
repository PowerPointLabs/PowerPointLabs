using System.Collections.Generic;
using PowerPointLabs.ImagesLab.Model;
using PowerPointLabs.ImagesLab.Thread.Interface;

namespace PowerPointLabs.ImagesLab.View.Interface
{
    public interface IImagesLabWindow
    {
        void ShowErrorMessageBox(string content);

        void ShowInfoMessageBox(string content);

        void ShowSuccessfullyAppliedDialog();

        void ActivateImageDownloadProgressRing();

        void DeactivateImageDownloadProgressRing();

        void UpdatePreviewImagesForDownloadedImage(ImageItem downloadedImageItem);

        IThreadContext GetThreadContext();

        string InitVariantsComboBox(IDictionary<string, List<StyleVariants>> variants);

        double GetVariationListBoxScrollOffset();

        void SetVariationListBoxScrollOffset(double offset);

        int GetVariationListBoxSelectedId();

        void SetVariationListBoxSelectedId(int index);

        string GetVariantsComboBoxSelectedValue();

        void UpdateStyleVariationsImages(IList<StyleOptions> givenOptions = null,
            Dictionary<string, List<StyleVariants>> givenVariants = null);
    }
}
