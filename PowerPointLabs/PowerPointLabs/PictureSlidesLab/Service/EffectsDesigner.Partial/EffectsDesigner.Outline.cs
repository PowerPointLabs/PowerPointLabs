using Microsoft.Office.Core;

using PowerPointLabs.PictureSlidesLab.Service.Effect;

using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointLabs.PictureSlidesLab.Service
{
    partial class EffectsDesigner
    {
        public PowerPoint.Shape ApplyRectOutlineEffect(PowerPoint.Shape imageShape, string overlayColor, int transparency)
        {
            TextBoxInfo tbInfo =
                new TextBoxes(Shapes.Range(), SlideWidth, SlideHeight)
                .GetTextBoxesInfo();
            if (tbInfo == null)
            {
                return null;
            }

            TextBoxes.AddMargin(tbInfo, 10);

            PowerPoint.Shape overlayShape = ApplyOverlayEffect(overlayColor, transparency, tbInfo.Left, tbInfo.Top, tbInfo.Width,
                tbInfo.Height);
            overlayShape.Fill.Visible = MsoTriState.msoFalse;
            overlayShape.Line.Visible = MsoTriState.msoTrue;

            ChangeName(overlayShape, EffectName.Banner);
            overlayShape.ZOrder(MsoZOrderCmd.msoSendToBack);
            if (imageShape != null)
            {
                imageShape.ZOrder(MsoZOrderCmd.msoSendToBack);
            }
            return overlayShape;
        }
    }
}
