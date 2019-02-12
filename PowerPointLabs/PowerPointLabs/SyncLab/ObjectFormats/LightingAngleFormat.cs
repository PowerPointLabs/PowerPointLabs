using System;
using System.Drawing;

using PowerPointLabs.ActionFramework.Common.Log;

using Shape = Microsoft.Office.Interop.PowerPoint.Shape;
using ThreeDFormat = Microsoft.Office.Interop.PowerPoint.ThreeDFormat;

namespace PowerPointLabs.SyncLab.ObjectFormats
{
    class LightingAngleFormat: Format
    {
        public override bool CanCopy(Shape formatShape)
        {
            return true;
        }

        public override void SyncFormat(Shape formatShape, Shape newShape)
        {
            if (!Sync(formatShape, newShape))
            {
                Logger.Log(newShape.Type + " unable to sync LightingEffect Format");
            }
        }

        public override Bitmap DisplayImage(Shape formatShape)
        {
            return SyncFormatUtil.GetTextDisplay(
                $"{formatShape.ThreeD.LightAngle}{SyncFormatConstants.DisplayDegreeSymbol}",
                SyncFormatConstants.DisplayImageFont,
                SyncFormatConstants.DisplayImageSize);
        }

        private static bool Sync(Shape formatShape, Shape newShape)
        {
            ThreeDFormat source = formatShape.ThreeD;
            ThreeDFormat dest = newShape.ThreeD;

            try
            {
                dest.LightAngle = source.LightAngle;
                return true;
            }
            catch (Exception e)
            {
                Logger.LogException(e, "Sync LightingAngleFormat");
                return false;
            }

        }
        

    }
}
