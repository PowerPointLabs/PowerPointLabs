using System;
using System.Drawing;

using Microsoft.Office.Core;

using PowerPointLabs.ActionFramework.Common.Log;

using Shape = Microsoft.Office.Interop.PowerPoint.Shape;

namespace PowerPointLabs.SyncLab.ObjectFormats
{
    class ReflectionDistanceFormat: Format
    {
        public override bool CanCopy(Shape formatShape)
        {
            return formatShape.Reflection.Type != MsoReflectionType.msoReflectionTypeNone;
        }

        public override void SyncFormat(Shape formatShape, Shape newShape)
        {
            if (!Sync(formatShape, newShape))
            {
                Logger.Log(newShape.Type + " unable to sync ReflectionDistance format");
            }
        }

        public override Bitmap DisplayImage(Shape formatShape)
        {
            return SyncFormatUtil.GetTextDisplay(
                Math.Round(formatShape.Reflection.Offset, 1).ToString(),
                SyncFormatConstants.DisplayImageFont,
                SyncFormatConstants.DisplayImageSize);
        }

        private static bool Sync(Shape formatShape, Shape newShape)
        {
            try
            {
                ReflectionFormat srcFormat = formatShape.Reflection;
                ReflectionFormat destFormat = newShape.Reflection;

                destFormat.Offset = srcFormat.Offset;
                return true;
            }
            catch (Exception e)
            {
                Logger.LogException(e, "Sync ReflectionDistanceFormat");
                return false;
            }
        }
    }
}
