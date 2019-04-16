using Microsoft.Office.Core;

using PowerPointLabs.PictureSlidesLab.Service.Effect;

using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointLabs.PictureSlidesLab.Service
{
    partial class EffectsDesigner
    {
        public PowerPoint.Shape ApplyRectBannerEffect(BannerDirection direction, Position textPos, PowerPoint.Shape imageShape,
            string overlayColor, int transparency)
        {
            TextBoxInfo tbInfo =
                new TextBoxes(Shapes.Range(), SlideWidth, SlideHeight)
                .GetTextBoxesInfo();
            if (tbInfo == null)
            {
                return null;
            }

            TextBoxes.AddMargin(tbInfo);

            PowerPoint.Shape overlayShape;
            direction = HandleAutoDirection(direction, textPos);
            switch (direction)
            {
                case BannerDirection.Horizontal:
                    overlayShape = ApplyOverlayEffect(overlayColor, transparency, 0, tbInfo.Top, SlideWidth,
                        tbInfo.Height);
                    break;
                // case BannerDirection.Vertical:
                default:
                    overlayShape = ApplyOverlayEffect(overlayColor, transparency, tbInfo.Left, 0, tbInfo.Width,
                        SlideHeight);
                    break;
            }
            ChangeName(overlayShape, EffectName.Banner);
            overlayShape.ZOrder(MsoZOrderCmd.msoSendToBack);
            if (imageShape != null)
            {
                imageShape.ZOrder(MsoZOrderCmd.msoSendToBack);
            }
            return overlayShape;
        }

        private BannerDirection HandleAutoDirection(BannerDirection dir, Position textPos)
        {
            if (dir != BannerDirection.Auto)
            {
                return dir;
            }

            switch (textPos)
            {
                case Position.Left:
                case Position.Centre:
                case Position.Right:
                    return BannerDirection.Vertical;
                default:
                    return BannerDirection.Horizontal;
            }
        }
    }
}
