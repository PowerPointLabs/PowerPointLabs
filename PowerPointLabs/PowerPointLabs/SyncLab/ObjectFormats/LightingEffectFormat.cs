using System;
using System.Drawing;

using Microsoft.Office.Core;
using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.ActionFramework.Common.Log;
using PowerPointLabs.Utils;

using Shape = Microsoft.Office.Interop.PowerPoint.Shape;
using Shapes = Microsoft.Office.Interop.PowerPoint.Shapes;
using ThreeDFormat = Microsoft.Office.Interop.PowerPoint.ThreeDFormat;

namespace PowerPointLabs.SyncLab.ObjectFormats
{
    class LightingEffectFormat: Format
    {
        public override bool CanCopy(Shape formatShape)
        {
            return true;
        }

        public override void SyncFormat(Shape formatShape, Shape newShape)
        {
            if (!Sync(formatShape, newShape))
            {
                Logger.Log(newShape.Type + " unable to sync LightingEffect Format");
            }
        }

        public override Bitmap DisplayImage(Shape formatShape)
        {
            Shapes shapes = SyncFormatUtil.GetTemplateShapes();
            Shape shape = shapes.AddShape(
                    MsoAutoShapeType.msoShapeRectangle, 0, 0, 
                    SyncFormatConstants.DisplayImageSize.Width,
                    SyncFormatConstants.DisplayImageSize.Height);
            shape.Line.Visible = MsoTriState.msoFalse;
            shape.ThreeD.Depth = SyncFormatConstants.DisplayImageDepth;
            shape.ThreeD.BevelTopType = SyncFormatConstants.DisplayBevelType;
            shape.ThreeD.BevelTopDepth = SyncFormatConstants.DisplayBevelHeight;
            shape.ThreeD.BevelTopInset = SyncFormatConstants.DisplayBevelWidth;
            shape.ThreeD.BevelBottomType = SyncFormatConstants.DisplayBevelType;
            shape.ThreeD.BevelBottomDepth = SyncFormatConstants.DisplayBevelHeight;
            shape.ThreeD.BevelBottomInset = SyncFormatConstants.DisplayBevelWidth;
            
            // setting mixed throws an exception
            if (formatShape.ThreeD.PresetLighting != MsoLightRigType.msoLightRigMixed)
            {
                shape.ThreeD.PresetLighting = formatShape.ThreeD.PresetLighting;
            }
            shape.ThreeD.SetPresetCamera(SyncFormatConstants.DisplayCameraPreset);
            
            Bitmap image = new Bitmap(GraphicsUtil.ShapeToBitmap(shape));
            shape.SafeDelete();
            return image;
        }

        private static bool Sync(Shape formatShape, Shape newShape)
        {
            ThreeDFormat source = formatShape.ThreeD;
            ThreeDFormat dest = newShape.ThreeD;

            try
            {
                // set lighting manually if type Mixed, setting type to Mixed throws an exception
                if (source.PresetLighting == MsoLightRigType.msoLightRigMixed)
                {
                    dest.PresetLightingDirection = source.PresetLightingDirection;
                    dest.PresetLightingSoftness = source.PresetLightingSoftness;
                }
                else
                {
                    // set lighting preset if not type Mixed
                    dest.PresetLighting = source.PresetLighting;
                }

                return true;
            }
            catch (Exception e)
            {
                Logger.LogException(e, "Sync LightingEffectFormat");
                return false;
            }

        }
        

    }
}
