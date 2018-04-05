using System.ComponentModel.Design;
using System.Drawing;
using Microsoft.Office.Core;
using PowerPointLabs.ActionFramework.Common.Log;
using PowerPointLabs.Utils;
using Shape = Microsoft.Office.Interop.PowerPoint.Shape;
using ThreeDFormat = Microsoft.Office.Interop.PowerPoint.ThreeDFormat;

namespace PowerPointLabs.SyncLab.ObjectFormats
{
    class BevelEffectFormat: Format
    {
        public override bool CanCopy(Shape formatShape)
        {
            return formatShape.ThreeD.BevelBottomType != MsoBevelType.msoBevelNone
                   || formatShape.ThreeD.BevelTopType != MsoBevelType.msoBevelNone;

        }

        public override void SyncFormat(Shape formatShape, Shape newShape)
        {
            if (!Sync(formatShape, newShape))
            {
                Logger.Log(newShape.Type + " unable to sync BevelEffect Format");
            }
        }

        public override Bitmap DisplayImage(Shape formatShape)
        {
            Bitmap image = GraphicsUtil.ShapeToBitmap(formatShape);
            return image;
        }

        private static bool Sync(Shape formatShape, Shape newShape)
        {
            ThreeDFormat source = formatShape.ThreeD;
            ThreeDFormat dest = newShape.ThreeD;

            try
            {
                // bottom bevel
                if (source.BevelBottomType != MsoBevelType.msoBevelTypeMixed)
                {
                    dest.BevelBottomType = source.BevelBottomType;
                }
                dest.BevelBottomDepth = source.BevelBottomDepth;
                dest.BevelBottomInset = source.BevelBottomInset;

                // top bevel
                if (source.BevelTopType != MsoBevelType.msoBevelTypeMixed)
                {
                    dest.BevelTopType = source.BevelTopType;
                }
                dest.BevelTopDepth = source.BevelTopDepth;
                dest.BevelTopInset = source.BevelTopInset;

                // contour 
                dest.ContourWidth = source.ContourWidth;
                // do not set SchemeColor, Brightness & ObjectThemeColor, setting them throws exceptions
                dest.ContourColor.RGB = source.ContourColor.RGB;
                dest.ContourColor.TintAndShade = source.ContourColor.TintAndShade;

                // depth (extrusion)
                dest.Depth = source.Depth;
                // don't sync type if type is TypeMixed, it throws an exception
                if (source.ExtrusionColorType != MsoExtrusionColorType.msoExtrusionColorTypeMixed)
                {
                    dest.ExtrusionColorType = source.ExtrusionColorType;
                }
                if (source.ExtrusionColorType != MsoExtrusionColorType.msoExtrusionColorAutomatic)
                {
                    // do not set SchemeColor, Brightness & ObjectThemeColor, setting them throws exceptions
                    dest.ExtrusionColor.ObjectThemeColor = source.ExtrusionColor.ObjectThemeColor;
                    dest.ExtrusionColor.RGB = source.ExtrusionColor.RGB;
                    dest.ExtrusionColor.TintAndShade = source.ExtrusionColor.TintAndShade;
                }

                // material
                dest.PresetMaterial = source.PresetMaterial;

                // lighting & angle
                dest.LightAngle = source.LightAngle;
                // set lighting to manually if of type Mixed
                if (source.PresetLighting == MsoLightRigType.msoLightRigMixed)
                {
                    dest.PresetLightingDirection = source.PresetLightingDirection;
                    dest.PresetLightingSoftness = source.PresetLightingSoftness;
                }
                else
                {
                    // set lighting preset if is not type Mixed
                    dest.PresetLighting = source.PresetLighting;
                }

                return true;
            }
            catch
            {
                return false;
            }

        }
        

    }
}
