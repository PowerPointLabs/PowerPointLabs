using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Drawing;
using System.Diagnostics;
using PowerPointLabs.Models;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;


namespace PowerPointLabs
{
    class PowerPointLabsGlobals
    {
        public static void Log(string logText, string type)
        {
            if (type.Equals("Info"))
                Trace.TraceInformation(DateTime.Now.ToString("yyyyMMddHHmmss") + ": " + logText);
            else if (type.Equals("Error"))
                Trace.TraceError(DateTime.Now.ToString("yyyyMMddHHmmss") + ": " + logText);
            else if (type.Equals("Warning"))
                Trace.TraceWarning(DateTime.Now.ToString("yyyyMMddHHmmss") + ": " + logText);
        }
        public static void LogException(Exception e, string methodName)
        {
            Log(methodName + ": " + e.Message + ": " + e.StackTrace, "Error");
        }

        public static void AddAckSlide()
        {
            try
            {
                PowerPointSlide lastSlide = PowerPointPresentation.Current.Slides.Last();
                if (!lastSlide.isAckSlide())
                    lastSlide.CreateAckSlide();
            }
            catch (Exception e)
            {
                PowerPointLabsGlobals.LogException(e, "AddAckSlide");
                throw;
            }
        }

        public static void FitShapeToSlide(ref PowerPoint.Shape shapeToMove)
        {
            shapeToMove.LockAspectRatio = Office.MsoTriState.msoFalse;
            shapeToMove.Left = 0;
            shapeToMove.Top = 0;
            shapeToMove.Width = PowerPointPresentation.Current.SlideWidth;
            shapeToMove.Height = PowerPointPresentation.Current.SlideHeight;
        }

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

        public static int CreateRGB(Color color)
        {
            // initial value
            int rgb = 0;

            // swap
            int red = color.B;
            int blue = color.R;
            int green = color.G;

            // create the newColor
            Color newColor = Color.FromArgb(red, green, blue);

            // set the return value
            rgb = newColor.ToArgb();

            // return value
            return rgb;
        }

        public static float GetMinimumRotation(float fromAngle, float toAngle)
        {
            fromAngle = Normalize(fromAngle);
            toAngle = Normalize(toAngle);

            float rotation1 = toAngle - fromAngle;
            float rotation2 = rotation1 == 0.0f ? 0.0f : Math.Abs(360.0f - Math.Abs(rotation1)) * (rotation1 / Math.Abs(rotation1)) * -1.0f;

            if (Math.Abs(rotation1) < Math.Abs(rotation2))
                return rotation1;
            else
                return rotation2;
        }

        public static float Normalize(float i)
        {
            //find effective angle
            float d = Math.Abs(i) % 360.0f;

            if (i < 0)
                return 360.0f - d; //return positive equivalent
            else
                return d;
        }
    }
}
