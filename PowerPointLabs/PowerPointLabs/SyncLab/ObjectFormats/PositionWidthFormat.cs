using System;
using System.Drawing;

using Microsoft.Office.Core;

using Shape = Microsoft.Office.Interop.PowerPoint.Shape;

namespace PowerPointLabs.SyncLab.ObjectFormats
{
    class PositionWidthFormat: Format
    {
        public override bool CanCopy(Shape formatShape)
        {
            return true;
        }

        public override void SyncFormat(Shape formatShape, Shape newShape)
        {
            MsoTriState lockState = newShape.LockAspectRatio;
            newShape.LockAspectRatio = MsoTriState.msoFalse;
            newShape.Width = formatShape.Width;
            newShape.LockAspectRatio = lockState;
        }

        public override Bitmap DisplayImage(Shape formatShape)
        {
            return SyncFormatUtil.GetTextDisplay(
                Math.Round(formatShape.Width).ToString(),
                SyncFormatConstants.DisplayImageFont,
                SyncFormatConstants.DisplayImageSize);
        }
    }
}
