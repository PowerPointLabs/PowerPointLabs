using System;
using System.ComponentModel.Design;
using System.Drawing;
using Microsoft.Office.Core;
using PowerPointLabs.ActionFramework.Common.Log;
using PowerPointLabs.Utils;
using Shape = Microsoft.Office.Interop.PowerPoint.Shape;
using ThreeDFormat = Microsoft.Office.Interop.PowerPoint.ThreeDFormat;

namespace PowerPointLabs.SyncLab.ObjectFormats
{
    class DepthSizeFormat: Format
    {
        public override bool CanCopy(Shape formatShape)
        {
            return true;
        }

        public override void SyncFormat(Shape formatShape, Shape newShape)
        {
            if (!Sync(formatShape, newShape))
            {
                Logger.Log(newShape.Type + " unable to sync DepthSize Format");
            }
        }

        public override Bitmap DisplayImage(Shape formatShape)
        {
            float depth = formatShape.ThreeD.Depth;
            if (IsFreshShape(formatShape))
            {
                // fresh shapes actually have 0 depth
                depth = 0f;
            }
            return SyncFormatUtil.GetTextDisplay(
                $"{Math.Round(depth, 1)} {SyncFormatConstants.DisplaySizeUnit}",
                SyncFormatConstants.DisplayImageFont,
                SyncFormatConstants.DisplayImageSize);
        }
        
        private static bool Sync(Shape formatShape, Shape newShape)
        {
            ThreeDFormat source = formatShape.ThreeD;
            ThreeDFormat dest = newShape.ThreeD;

            try
            {
                if (IsFreshShape(formatShape) && IsFreshShape(newShape))
                {
                    // both are fresh shapes, do nothing
                    return true;
                }

                float depth = source.Depth;
                if (IsFreshShape(formatShape))
                {
                    // fresh shapes actually have 0 depth
                    depth = 0f;
                }

                dest.Depth = depth;
                
                return true;
            }
            catch
            {
                return false;
            }

        }
        
        /**
         * Checks if a shape does not have any 3d setting edited (excluding bevel)
         * The API gives the wrong depth value for these shapes, it should return 0 instead of 36.
         */
        private static bool IsFreshShape(Shape shape)
        {
            // PresetMaterial == Mixed and LightRig == Mixed when a shape is just created
            // and 3d effects have not been modified
            bool isFreshShape = shape.ThreeD.PresetMaterial == MsoPresetMaterial.msoPresetMaterialMixed
                                && shape.ThreeD.PresetLighting == MsoLightRigType.msoLightRigMixed;
            return isFreshShape;
        }

    }
}
