using System;
using System.ComponentModel.Design;
using System.Drawing;
using System.Text.RegularExpressions;
using Microsoft.Office.Core;
using PowerPointLabs.ActionFramework.Common.Log;
using PowerPointLabs.Utils;
using Shape = Microsoft.Office.Interop.PowerPoint.Shape;
using ThreeDFormat = Microsoft.Office.Interop.PowerPoint.ThreeDFormat;

namespace PowerPointLabs.SyncLab.ObjectFormats
{
    class LightingAngleFormat: Format
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
            return SyncFormatUtil.GetTextDisplay(
                $"{formatShape.ThreeD.LightAngle}{SyncFormatConstants.DisplayDegreeSymbol}",
                SyncFormatConstants.DisplayImageFont,
                SyncFormatConstants.DisplayImageSize);
        }

        private static bool Sync(Shape formatShape, Shape newShape)
        {
            ThreeDFormat source = formatShape.ThreeD;
            ThreeDFormat dest = newShape.ThreeD;

            try
            {
                // set lighting manually if of type Mixed
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
