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

            // make reflection more visible in preview
            // changing any Reflection setting sets ReflectionFormat.Type to msoReflectionTypeMixed 
            // perform configuration on a duplicate to avoid complex control flow required to
            // restore ReflectionFormat.Type
            Shape duplicate = formatShape.Duplicate()[1];
            duplicate.Reflection.Size = 100.0f;
            duplicate.Reflection.Transparency = 0.1f;
            duplicate.Reflection.Blur = 0.5f;
            
            Bitmap image = Graphics.ShapeToBitmap(formatShape);
            duplicate.Delete();

            return image;
        }

        private static bool Sync(Shape formatShape, Shape newShape)
        {
            try
            {
                ReflectionFormat srcFormat = formatShape.Reflection;
                ReflectionFormat destFormat = newShape.Reflection;

                if (srcFormat.Type != MsoReflectionType.msoReflectionTypeMixed)
                {
                    // setting ReflectionFormat.Type automatically sets Reflection settings
                    // there is no need to set them manually
                    destFormat.Type = srcFormat.Type;
                }
                else
                {
                    // setting mixed type throws an exception
                    // skip it and set Reflection settings manually
                    destFormat.Blur = srcFormat.Blur;
                    destFormat.Offset = srcFormat.Offset;
                    destFormat.Size = srcFormat.Size;
                    destFormat.Transparency = srcFormat.Transparency;
                }
            }
            catch (Exception)
            {
                return false;
            }
            return true;
        }
    }
}
