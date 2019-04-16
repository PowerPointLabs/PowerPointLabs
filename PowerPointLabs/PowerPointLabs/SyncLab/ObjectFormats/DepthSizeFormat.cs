using System;
using System.Drawing;

using Microsoft.Office.Core;

using PowerPointLabs.ActionFramework.Common.Log;

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
            if (HasErrorneousDepth(formatShape))
            {
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
                if (HasErrorneousDepth(formatShape) && HasErrorneousDepth(newShape))
                {
                    // both have no 3d settings, do nothing
                    // setting depth here changes the color of the shape slightly for unknown reasons
                    // we cannot revert this change
                    return true;
                }

                float depth = source.Depth;
                if (HasErrorneousDepth(formatShape))
                {
                    // fresh shapes actually have 0 depth
                    depth = 0f;
                }

                dest.Depth = depth;
                
                return true;
            }
            catch (Exception e)
            {
                Logger.LogException(e, "Sync DepthSizeFormat");
                return false;
            }

        }
        
        /**
         * Checks if the shape gives the wrong depth
         * The API gives the wrong depth value at times (gives 36, when it should be 0)
         */
        private static bool HasErrorneousDepth(Shape shape)
        {
            // seems to happen PresetMaterial == Mixed 
            // the Mixed type seems to be reserved for shapes with untouched depth, material or contour
            return shape.ThreeD.PresetMaterial == MsoPresetMaterial.msoPresetMaterialMixed;
        }

    }
}
