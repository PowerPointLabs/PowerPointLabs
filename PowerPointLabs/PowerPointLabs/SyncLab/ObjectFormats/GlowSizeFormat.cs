﻿using System;
using System.Drawing;

using Microsoft.Office.Core;

using PowerPointLabs.ActionFramework.Common.Log;

using Shape = Microsoft.Office.Interop.PowerPoint.Shape;

namespace PowerPointLabs.SyncLab.ObjectFormats
{
    class GlowSizeFormat: Format
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
                Logger.Log(newShape.Type + " unable to sync GlowSize Format");
            }
        }

        public override Bitmap DisplayImage(Shape formatShape)
        {
            return SyncFormatUtil.GetTextDisplay(
                Math.Round(formatShape.Glow.Radius).ToString(),
                SyncFormatConstants.DisplayImageFont,
                SyncFormatConstants.DisplayImageSize);
        }

        private static bool Sync(Shape formatShape, Shape newShape)
        {
            try
            {
                GlowFormat dest = newShape.Glow;
                GlowFormat source = formatShape.Glow;

                dest.Radius = source.Radius;
                return true;
            }
            catch (Exception e)
            {
                Logger.LogException(e, "Sync GlowSizeFormat");
                return false;
            }
        }
    }
}
