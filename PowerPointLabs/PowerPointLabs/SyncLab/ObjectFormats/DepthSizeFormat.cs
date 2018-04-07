using System;
using System.ComponentModel.Design;
using System.Drawing;
using Microsoft.Office.Core;
using PowerPointLabs.ActionFramework.Common.Log;
using PowerPointLabs.Utils;
using Shape = Microsoft.Office.Interop.PowerPoint.Shape;
using ThreeDFormat = Microsoft.Office.Interop.PowerPoint.ThreeDFormat;

namespace PowerPointLabs.SyncLab.ObjectFormats
{
    class DepthSizeFormat: Format
    {
        public override bool CanCopy(Shape formatShape)
        {
            return true;
        }

        public override void SyncFormat(Shape formatShape, Shape newShape)
        {
            if (!Sync(formatShape, newShape))
            {
                Logger.Log(newShape.Type + " unable to sync DepthSize Format");
            }
        }

        public override Bitmap DisplayImage(Shape formatShape)
        {
            return SyncFormatUtil.GetTextDisplay(
                $"{Math.Round(formatShape.ThreeD.Depth)} {SyncFormatConstants.DisplaySizeUnit}",
                SyncFormatConstants.DisplayImageFont,
                SyncFormatConstants.DisplayImageSize);
        }
        
        private static bool Sync(Shape formatShape, Shape newShape)
        {
            ThreeDFormat source = formatShape.ThreeD;
            ThreeDFormat dest = newShape.ThreeD;

            try
            {
                dest.Depth = source.Depth;
                return true;
            }
            catch
            {
                return false;
            }

        }
        

    }
}
