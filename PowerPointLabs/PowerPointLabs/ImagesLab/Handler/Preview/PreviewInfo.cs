using PowerPointLabs.ImagesLab.Util;

namespace PowerPointLabs.ImagesLab.Handler.Preview
{
    public class PreviewInfo
    {
        public string OverlayStyleImagePath { get; private set; }
        public string SpecialEffectStyleImagePath { get; private set; }
        public string BannerStyleImagePath { get; private set; }
        public string TextboxStyleImagePath { get; private set; }
        public string BlurStyleImagePath { get; private set; }
        public string DirectTextStyleImagePath { get; private set; }

        public string PreviewApplyStyleImagePath { get; private set; }

        public PreviewInfo()
        {
            Init();
        }

        private void Init()
        {
            TextboxStyleImagePath = TempPath.GetPath("textbox");
            BlurStyleImagePath = TempPath.GetPath("blur");
            DirectTextStyleImagePath = TempPath.GetPath("directtext");
            BannerStyleImagePath = TempPath.GetPath("banner");
            SpecialEffectStyleImagePath = TempPath.GetPath("specialeffect");
            OverlayStyleImagePath = TempPath.GetPath("overlay");

            PreviewApplyStyleImagePath = TempPath.GetPath("previewapply");
        }
    }
}
