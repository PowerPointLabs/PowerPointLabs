using System.ComponentModel;

using PowerPointLabs.PictureSlidesLab.Service.Effect;

namespace PowerPointLabs.PictureSlidesLab.Model
{
    partial class StyleOption
    {
        #region APIs
        public string GetFontFamily()
        {
            return FontFamily;
        }

        public Position GetTextBoxPosition()
        {
            switch (TextBoxPosition)
            {
                case 0:
                    return Position.NoEffect;
                case 1:
                    return Position.TopLeft;
                case 2:
                    return Position.Top;
                case 3:
                    return Position.TopRight;
                case 4:
                    return Position.Left;
                case 5:
                    return Position.Centre;
                case 6:
                    return Position.Right;
                case 7:
                    return Position.BottomLeft;
                case 8:
                    return Position.Bottom;
                case 9:
                    return Position.BottomRight;
                default:
                    return Position.NoEffect;
            }
        }

        public Alignment GetTextAlignment()
        {
            switch (TextBoxAlignment)
            {
                case 0:
                    return Alignment.Auto;
                case 1:
                    return Alignment.Left;
                case 2:
                    return Alignment.Centre;
                case 3:
                    return Alignment.Right;
                default:
                    return Alignment.NoEffect;
            }
        }
        #endregion

        [DefaultValue(true)]
        public bool IsUseTextFormat { get; set; }

        [DefaultValue("Calibri")]
        public string FontFamily { get; set; }

        [DefaultValue(0)]
        public int FontSizeIncrease { get; set; }

        [DefaultValue("#FFFFFF")]
        public string FontColor { get; set; }

        [DefaultValue(0)]
        public int TextTransparency { get; set; }

        [DefaultValue(5)]
        public int TextBoxPosition { get; set; }

        [DefaultValue(0)]
        public int TextBoxAlignment { get; set; }
    }
}
