using PowerPointLabs.PictureSlidesLab.Util;

namespace PowerPointLabs.PictureSlidesLab.Service.Preview
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
