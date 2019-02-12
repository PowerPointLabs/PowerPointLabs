using System;
using System.Drawing;

using Microsoft.Office.Core;

using PowerPointLabs.ActionFramework.Common.Log;

using Shape = Microsoft.Office.Interop.PowerPoint.Shape;

namespace PowerPointLabs.SyncLab.ObjectFormats
{
    class ReflectionTransparencyFormat: Format
    {
        public override bool CanCopy(Shape formatShape)
        {
            return formatShape.Reflection.Type != MsoReflectionType.msoReflectionTypeNone;
        }

        public override void SyncFormat(Shape formatShape, Shape newShape)
        {
            if (!Sync(formatShape, newShape))
            {
                Logger.Log(newShape.Type + " unable to sync ReflectionTransparency format");
            }
        }

        public override Bitmap DisplayImage(Shape formatShape)
        {
            // transparency is a float from 0-1, multiply by 100 to get percentage
            return SyncFormatUtil.GetTextDisplay(
                Math.Round(formatShape.Reflection.Transparency * 100).ToString() + "%",
                SyncFormatConstants.DisplayImageFont,
                SyncFormatConstants.DisplayImageSize);
        }

        private static bool Sync(Shape formatShape, Shape newShape)
        {
            try
            {
                ReflectionFormat srcFormat = formatShape.Reflection;
                ReflectionFormat destFormat = newShape.Reflection;

                destFormat.Transparency = srcFormat.Transparency;
                return true;
            }
            catch (Exception e)
            {
                Logger.LogException(e, "Sync ReflectionTransparencyFormat");
                return false;
            }
        }
    }
}
