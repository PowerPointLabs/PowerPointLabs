using PowerPointLabs.ImagesLab.Util;

namespace PowerPointLabs.ImagesLab.Handler.Preview
{
    public class PreviewInfo
    {
        public string PreviewApplyStyleImagePath { get; private set; }

        public PreviewInfo()
        {
            PreviewApplyStyleImagePath = TempPath.GetPath("previewapply");
        }
    }
}
