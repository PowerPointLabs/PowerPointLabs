using System;

using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;


namespace PowerPointLabs
{
    public class LegacyShapeUtil
    {
        // TODO: the following three functions must be renamed, the function name does not match what the function does
        public static void CopyShapePosition(PowerPoint.Shape shapeToCopy, ref PowerPoint.Shape shapeToMove)
        {
            shapeToMove.Left = shapeToCopy.Left + (shapeToCopy.Width / 2) - (shapeToMove.Width / 2);
            shapeToMove.Top = shapeToCopy.Top + (shapeToCopy.Height / 2) - (shapeToMove.Height / 2);
        }

        public static void CopyShapeSize(PowerPoint.Shape shapeToCopy, ref PowerPoint.Shape shapeToMove)
        {
            shapeToMove.LockAspectRatio = Office.MsoTriState.msoFalse;
            shapeToMove.Width = shapeToCopy.Width;
            shapeToMove.Height = shapeToCopy.Height;
        }

        public static void CopyShapeAttributes(PowerPoint.Shape shapeToCopy, ref PowerPoint.Shape shapeToMove)
        {
            CopyShapeSize(shapeToCopy, ref shapeToMove);
            CopyShapePosition(shapeToCopy, ref shapeToMove);
        }

        public static float GetMinimumRotation(float fromAngle, float toAngle)
        {
            fromAngle = Normalize(fromAngle);
            toAngle = Normalize(toAngle);

            float rotation1 = toAngle - fromAngle;
            float rotation2 = rotation1 == 0.0f ? 0.0f : Math.Abs(360.0f - Math.Abs(rotation1)) * (rotation1 / Math.Abs(rotation1)) * -1.0f;

            if (Math.Abs(rotation1) < Math.Abs(rotation2))
            {
                return rotation1;
            }
            else
            {
                return rotation2;
            }
        }

        private static float Normalize(float i)
        {
            //find effective angle
            float d = Math.Abs(i) % 360.0f;

            if (i < 0)
            {
                return 360.0f - d; //return positive equivalent
            }
            else
            {
                return d;
            }
        }
    }
}
