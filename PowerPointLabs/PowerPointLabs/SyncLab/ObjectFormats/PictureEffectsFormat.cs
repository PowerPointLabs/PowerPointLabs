using System;
using System.Drawing;
using Microsoft.Office.Core;
using PowerPointLabs.ActionFramework.Common.Log;
using PowerPointLabs.Utils;
using Shape = Microsoft.Office.Interop.PowerPoint.Shape;
using Shapes = Microsoft.Office.Interop.PowerPoint.Shapes;

namespace PowerPointLabs.SyncLab.ObjectFormats
{
    class PictureEffectsFormat
    {
        public static bool CanCopy(Shape formatShape)
        {
            try
            {
                return formatShape.Fill.PictureEffects.Count > 0;
            } 
            catch
            {
                return false;
            }
        }

        public static void SyncFormat(Shape formatShape, Shape newShape)
        {
            if (Sync(formatShape, newShape))
            {
                Logger.Log(newShape.Type + " unable to sync Font Format");
            }
        }

        public static Bitmap DisplayImage(Shape formatShape)
        {
            Bitmap image = GraphicsUtil.ShapeToBitmap(formatShape);
            return image;
        }

        private static bool Sync(Shape formatShape, Shape newShape)
        {
            try
            {
                PictureEffects dest = newShape.Fill.PictureEffects;
                PictureEffects source = formatShape.Fill.PictureEffects;
                
                // clear current effects
                for (int i = 1; i <= dest.Count; i++)
                {
                    dest.Delete(i);
                }

                // add new effects
                for (int i = 1; i <= source.Count; i++)
                {
                    PictureEffect effect = source[i];
                    dest.Insert(effect.Type, i);
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
