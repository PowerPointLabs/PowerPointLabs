using System.ComponentModel;

using PowerPointLabs.PictureSlidesLab.Service.Effect;

namespace PowerPointLabs.PictureSlidesLab.Model
{
    partial class StyleOption
    {
        #region APIs
        public BannerShape GetBannerShape()
        {
            switch (BannerShape)
            {
                case 0:
                    return Service.Effect.BannerShape.Rectangle;
                case 1:
                    return Service.Effect.BannerShape.Circle;
                case 2:
                    return Service.Effect.BannerShape.RectangleOutline;
                default:
                    return Service.Effect.BannerShape.CircleOutline;
            }
        }

        public BannerDirection GetBannerDirection()
        {
            switch (BannerDirection)
            {
                case 0:
                    return Service.Effect.BannerDirection.Auto;
                case 1:
                    return Service.Effect.BannerDirection.Horizontal;
                // case 2:
                default:
                    return Service.Effect.BannerDirection.Vertical;
            }
        }
        #endregion

        [DefaultValue(false)]
        public bool IsUseBannerStyle { get; set; }

        [DefaultValue(0)]
        public int BannerShape { get; set; }

        [DefaultValue(0)]
        public int BannerDirection { get; set; }

        [DefaultValue("#000000")]
        public string BannerColor { get; set; }

        [DefaultValue(25)]
        public int BannerTransparency { get; set; }
    }
}
