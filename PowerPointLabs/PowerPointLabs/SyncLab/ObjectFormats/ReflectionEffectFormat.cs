using System;
using System.Drawing;
using Microsoft.Office.Core;
using PowerPointLabs.ActionFramework.Common.Log;
using Graphics = PowerPointLabs.Utils.GraphicsUtil;
using Shape = Microsoft.Office.Interop.PowerPoint.Shape;

namespace PowerPointLabs.SyncLab.ObjectFormats
{
    class ReflectionEffectFormat: Format
    {
        public override bool CanCopy(Shape formatShape)
        {
            return formatShape.Reflection.Type != MsoReflectionType.msoReflectionTypeNone;
        }

        public override void SyncFormat(Shape formatShape, Shape newShape)
        {
            if (!Sync(formatShape, newShape))
            {
                Logger.Log(newShape.Type + " unable to sync Reflection");
            }
        }

        public override Bitmap DisplayImage(Shape formatShape)
        {

            // make reflection more visible 
            float oldTransparency = formatShape.Reflection.Transparency;
            float oldSize = formatShape.Reflection.Size;
            float oldBlur = formatShape.Reflection.Blur;
            formatShape.Reflection.Size = 100.0f;
            formatShape.Reflection.Transparency = 0.1f;
            formatShape.Reflection.Blur = 0.5f;
            Bitmap image = Graphics.ShapeToBitmap(formatShape);

            formatShape.Reflection.Size = oldSize;
            formatShape.Reflection.Transparency = oldTransparency;
            formatShape.Reflection.Blur = oldBlur;

            return image;
        }

        private static bool Sync(Shape formatShape, Shape newShape)
        {
            try
            {
                ReflectionFormat srcFormat = formatShape.Reflection;
                ReflectionFormat destFormat = newShape.Reflection;

                destFormat.Type = srcFormat.Type;
                destFormat.Blur = srcFormat.Blur;
                destFormat.Offset = srcFormat.Offset;
                destFormat.Size = srcFormat.Size;
                destFormat.Transparency = srcFormat.Transparency;
            }
            catch (Exception)
            {
                return false;
            }
            return true;
        }
    }
}
