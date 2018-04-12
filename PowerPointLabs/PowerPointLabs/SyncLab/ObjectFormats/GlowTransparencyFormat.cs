using System;
using System.Drawing;
using Microsoft.Office.Core;
using PowerPointLabs.ActionFramework.Common.Log;
using PowerPointLabs.Utils;
using Shape = Microsoft.Office.Interop.PowerPoint.Shape;
using Shapes = Microsoft.Office.Interop.PowerPoint.Shapes;

namespace PowerPointLabs.SyncLab.ObjectFormats
{
    class GlowTransparencyFormat: Format
    {
        public override bool CanCopy(Shape formatShape)
        {
            GlowFormat glow = formatShape.Glow;
            return glow.Radius > 0 && glow.Transparency > 0.0f;
        }

        public override void SyncFormat(Shape formatShape, Shape newShape)
        {
            if (!Sync(formatShape, newShape))
            {
                Logger.Log(newShape.Type + " unable to sync GlowTransparency Format");
            }
        }

        public override Bitmap DisplayImage(Shape formatShape)
        {
            return SyncFormatUtil.GetTextDisplay(
                Math.Round(formatShape.Glow.Transparency * 100).ToString() + "%",
                SyncFormatConstants.DisplayImageFont,
                SyncFormatConstants.DisplayImageSize);
        }

        private static bool Sync(Shape formatShape, Shape newShape)
        {
            try
            {
                GlowFormat dest = newShape.Glow;
                GlowFormat source = formatShape.Glow;

                dest.Transparency = source.Transparency;
            }
            catch (Exception)
            {
                return false;
            }
            return true;
        }
    }
}
