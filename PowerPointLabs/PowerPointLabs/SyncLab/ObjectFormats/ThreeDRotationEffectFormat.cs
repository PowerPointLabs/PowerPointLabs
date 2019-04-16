using System;
using System.Drawing;

using Microsoft.Office.Core;

using PowerPointLabs.ActionFramework.Common.Log;
using PowerPointLabs.Utils;

using Shape = Microsoft.Office.Interop.PowerPoint.Shape;
using ThreeDFormat = Microsoft.Office.Interop.PowerPoint.ThreeDFormat;

namespace PowerPointLabs.SyncLab.ObjectFormats
{
    class ThreeDRotationEffectFormat: Format
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
                Logger.Log(newShape.Type + " unable to sync 3DRotation Format");
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
                if (source.PresetThreeDFormat == MsoPresetThreeDFormat.msoPresetThreeDFormatMixed)
                {
                    // skip setting PresetCamera if Mixed. setting it throws an error
                    if (source.PresetCamera != MsoPresetCamera.msoPresetCameraMixed)
                    {
                        dest.SetPresetCamera(source.PresetCamera);
                    }
                    // set FieldOfView anyway, PresetCamera doesn't seem to set it
                    dest.FieldOfView = source.FieldOfView;

                    dest.RotationX = source.RotationX;
                    dest.RotationZ = source.RotationZ;
                    dest.RotationY = source.RotationY;

                    // set perspective only if it is different,
                    // setting the same perspective applies an unknown change to the lighting of the shape
                    // this change is visible to the eye, but we cannot undo it
                    if (dest.Perspective != source.Perspective)
                    {
                        dest.Perspective = source.Perspective;
                    }
                }
                else
                {
                    dest.SetThreeDFormat(source.PresetThreeDFormat);
                }


                // ThreeDFormat.Z must be between -4000 & 4000 exclusive.
                // when source.Z > 4000 or source.Z < - 4000, it actually means 0
                float nearestZ = source.Z;
                nearestZ = nearestZ > 4000 ? 0f : nearestZ;
                nearestZ = nearestZ < -4000 ? 0f : nearestZ;
                dest.Z = nearestZ;
                
                dest.ProjectText = source.ProjectText;
                return true;
            }
            catch (Exception e)
            {
                Logger.LogException(e, "Sync ThreeDRotationEffectFormat");
                return false;
            }
            
        }
    }
}
