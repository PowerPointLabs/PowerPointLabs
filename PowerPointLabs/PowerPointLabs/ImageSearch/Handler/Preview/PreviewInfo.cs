using PowerPointLabs.ImageSearch.Util;

namespace PowerPointLabs.ImageSearch.Handler.Preview
{
    public class PreviewInfo
    {
        public string SpecialEffectStyleImagePath { get; private set; }
        public string TextboxStyleImagePath { get; private set; }
        public string BlurStyleImagePath { get; private set; }
        public string DirectTextStyleImagePath { get; private set; }

        public PreviewInfo()
        {
            Init();
        }

        private void Init()
        {
            TextboxStyleImagePath = TempPath.GetPath("textbox");
            BlurStyleImagePath = TempPath.GetPath("blur");
            DirectTextStyleImagePath = TempPath.GetPath("directtext");
            SpecialEffectStyleImagePath = TempPath.GetPath("specialeffect");
        }
    }
}
