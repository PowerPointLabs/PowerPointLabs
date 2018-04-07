using System;
using System.Collections.Generic;
using System.Drawing;
using Microsoft.Office.Core;
using PowerPointLabs.ActionFramework.Common.Log;
using PowerPointLabs.Utils;
using Shape = Microsoft.Office.Interop.PowerPoint.Shape;
using Shapes = Microsoft.Office.Interop.PowerPoint.Shapes;

namespace PowerPointLabs.SyncLab.ObjectFormats
{
    class ArtisticEffectFormat: Format
    {
        public override bool CanCopy(Shape formatShape)
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

        public override void SyncFormat(Shape formatShape, Shape newShape)
        {
            if (!Sync(formatShape, newShape))
            {
                Logger.Log(newShape.Type + " unable to sync ArtisticEffect Format");
            }
        }

        public override Bitmap DisplayImage(Shape formatShape)
        {
            Bitmap image = GraphicsUtil.ShapeToBitmap(formatShape);
            return image;
        }

        public static List<MsoPictureEffectType> GetArtisticEffects(Shape shape)
        {
            List<MsoPictureEffectType> artisticEffects = new List<MsoPictureEffectType>();
            try
            {
                PictureEffects effects = shape.Fill.PictureEffects;
                for (int i = 1; i <= effects.Count; i++)
                {
                    PictureEffect effect = effects[i];
                    artisticEffects.Add(effect.Type);
                }
            }
            catch (Exception)
            {
                // do nothing, shape does not support picture effect
            }

            return artisticEffects;
        }

        public static void ClearArtisticEffects(Shape shape)
        {
            try
            {
                PictureEffects dest = shape.Fill.PictureEffects;
                for (int i = 1; i <= dest.Count; i++)
                {
                    dest.Delete(i);
                }
            }
            catch (Exception)
            {
                // ignore the exception, this shape is not compatible with ArtisticEffects.
            }
        }

        public static void ApplyArtisticEffects(Shape shape, List<MsoPictureEffectType> effectTypes)
        {
            // add new effects
            try
            {
                PictureEffects dest = shape.Fill.PictureEffects;
                for (int i = 0; i < effectTypes.Count; i++)
                {
                    int index = i + 1;
                    dest.Insert(effectTypes[i], index);
                }
            }
            catch (Exception)
            {
                // ignore the exception, this shape is not compatible with ArtisticEffects.
            }
        }
        
        private static bool Sync(Shape formatShape, Shape newShape)
        {
            try
            {
                // access PictureEffects, just to make sure shapes are compatible with ArtisticEffect
                // will throw an exception otherwise
                PictureEffects dest = newShape.Fill.PictureEffects;
                PictureEffects source = formatShape.Fill.PictureEffects;
                
                List<MsoPictureEffectType> effectTypes = GetArtisticEffects(formatShape);
                ClearArtisticEffects(newShape);
                ApplyArtisticEffects(newShape, effectTypes);
            }
            catch (Exception)
            {
                return false;
            }
            return true;
        }
    }
}
