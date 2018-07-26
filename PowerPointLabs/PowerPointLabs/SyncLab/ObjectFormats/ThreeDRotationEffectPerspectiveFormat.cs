using System;
using System.CodeDom;
using System.ComponentModel.Design;
using System.Drawing;
using System.Windows;
using Microsoft.Office.Core;
using PowerPointLabs.ActionFramework.Common.Log;
using PowerPointLabs.EffectsLab;
using PowerPointLabs.Utils;
using Shape = Microsoft.Office.Interop.PowerPoint.Shape;
using Shapes = Microsoft.Office.Interop.PowerPoint.Shapes;
using ThreeDFormat = Microsoft.Office.Interop.PowerPoint.ThreeDFormat;

namespace PowerPointLabs.SyncLab.ObjectFormats
{
    class ThreeDRotationEffectPerspectiveFormat : Format
    {
        private static readonly float TOLERANCE = Single.Epsilon;

        public override bool CanCopy(Shape formatShape)
        {
            ThreeDFormat threeD = formatShape.ThreeD;

            // equality check for floating point numbers
            return Math.Abs(threeD.RotationX) > TOLERANCE
                   || Math.Abs(threeD.RotationY) > TOLERANCE
                   || Math.Abs(threeD.RotationZ) > TOLERANCE
                   || Math.Abs(threeD.FieldOfView) > TOLERANCE
                   || threeD.Perspective == MsoTriState.msoTrue;
        }

        public override void SyncFormat(Shape formatShape, Shape newShape)
        {
            if (!Sync(formatShape, newShape))
            {
                Logger.Log(newShape.Type + " unable to sync " + this.GetType().Name);
            }
        }

        public override Bitmap DisplayImage(Shape formatShape)
        {
            return SyncFormatUtil.GetTextDisplay(
                ((formatShape.ThreeD.ProjectText == MsoTriState.msoTrue) ? "Yes" : "No"),
                SyncFormatConstants.DisplayImageFont,
                SyncFormatConstants.DisplayImageSize);
        }

        private static bool Sync(Shape formatShape, Shape newShape)
        {
            ThreeDFormat source = formatShape.ThreeD;
            ThreeDFormat dest = newShape.ThreeD;
            
            try
            {
                // set perspective only if it is different,
                // setting the same perspective applies an unknown change to the lighting of the shape
                // this change is visible to the eye, but we cannot undo it
                if (dest.Perspective != source.Perspective)
                {
                    dest.Perspective = source.Perspective;
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
