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
    class BevelTopFormat: Format
    {
        public override bool CanCopy(Shape formatShape)
        {
            return true;
        }

        public override void SyncFormat(Shape formatShape, Shape newShape)
        {
            if (!Sync(formatShape, newShape))
            {
                Logger.Log(newShape.Type + " unable to sync BevelTop Format");
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
            
            // don't set type if type is TypeMixed, it throws an exception
            if (formatShape.ThreeD.BevelTopType != MsoBevelType.msoBevelTypeMixed)
            {
                shape.ThreeD.BevelTopType = formatShape.ThreeD.BevelTopType;
                // set depth & inset only if type is not none,
                // adjusting these 2 will automatically set type from None to Round
                if (shape.ThreeD.BevelTopType != MsoBevelType.msoBevelNone)
                {
                    shape.ThreeD.BevelTopDepth = SyncFormatConstants.DisplayBevelHeight;
                    shape.ThreeD.BevelTopInset = SyncFormatConstants.DisplayBevelWidth;
                }
            }
            shape.ThreeD.BevelBottomType = MsoBevelType.msoBevelNone;
            shape.ThreeD.SetPresetCamera(MsoPresetCamera.msoCameraPerspectiveAbove);
            shape.ThreeD.PresetLighting = MsoLightRigType.msoLightRigBalanced;
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
                // don't set type if type is TypeMixed, it throws an exception
                if (source.BevelTopType != MsoBevelType.msoBevelTypeMixed)
                {
                    dest.BevelTopType = source.BevelTopType;
                    // set depth & inset only if type is not none,
                    // adjusting these 2 will automatically set type from None to Round
                    if (source.BevelTopType != MsoBevelType.msoBevelNone)
                    {
                        // set the settings anyway, setting the type alone is insufficient
                        dest.BevelTopDepth = source.BevelTopDepth;
                        dest.BevelTopInset = source.BevelTopInset;
                    }
                }

                return true;
            }
            catch (Exception e)
            {
                Logger.LogException(e, "Sync BevelTopFormat");
                return false;
            }

        }
        

    }
}
