using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using PowerPointLabs.Models;
using Shape = Microsoft.Office.Interop.PowerPoint.Shape;


namespace PowerPointLabs.ResizeLab
{
    /// <summary>
    /// Handles the stretching and shrinking of shapes in resize lab
    /// </summary>
    internal partial class ResizeLabMain
    {

        private const int ModShapesIndex = 2;
        /// <summary>
        /// Stretches a given shape to match an edge of the reference shape.
        /// </summary>
        /// <param name="referenceShape">The edge to refer to</param>
        /// <param name="stretchShape">The shape to stretch</param>
        /// <returns>True if shape is stretched successfully, false otherwise</returns>
        private delegate bool StretchAction(PowerPoint.Shape referenceShape, PowerPoint.Shape stretchShape);
        #region API

        /// <summary>
        /// Stretches all selected shapes to the left edge of reference shape
        /// </summary>
        /// <param name="stretchShapes">The shapes to stretch</param>
        public void StretchLeft(PowerPoint.ShapeRange stretchShapes)
        {

            var sa = new StretchAction((PowerPoint.Shape referenceShape, PowerPoint.Shape stretchShape) =>
            {
                // Stretching the shape will cause the object to be very small
                if (GetRight(stretchShape) < referenceShape.Left)
                {
                    return false;
                }
                // The actual stretch action
                stretchShape.Width += stretchShape.Left - referenceShape.Left;
                stretchShape.Left = referenceShape.Left;

                return true;
            });
            Stretch(stretchShapes, sa);
        }

        /// <summary>
        /// Stretches all selected shapes to the right edge of reference shape
        /// </summary>
        /// <param name="stretchShapes">The shapes to stretch</param>
        public void StretchRight(PowerPoint.ShapeRange stretchShapes)
        {
            var sa = new StretchAction((PowerPoint.Shape referenceShape, PowerPoint.Shape stretchShape) =>
            {
                // Stretching the shape will cause the object to be very small
                if (stretchShape.Left > GetRight(referenceShape))
                {
                    return false;
                }
                // The actual stretch action
                stretchShape.Width += GetRight(referenceShape) - GetRight(stretchShape);

                return true;
            });
            Stretch(stretchShapes, sa);
        }

        /// <summary>
        /// Stretches all selected shapes to the top edge of reference shape
        /// </summary>
        /// <param name="stretchShapes">The shapes to stretch</param>
        public void StretchTop(PowerPoint.ShapeRange stretchShapes)
        {
            var sa = new StretchAction((PowerPoint.Shape referenceShape, PowerPoint.Shape stretchShape) =>
            {
                // Stretching the shape will cause the object to be very small
                if (GetBottom(stretchShape) < referenceShape.Top)
                {
                    return false;
                }
                // The actual stretch action
                stretchShape.Height += stretchShape.Top - referenceShape.Top;
                stretchShape.Top = referenceShape.Top;

                return true;
            });
            Stretch(stretchShapes, sa);
        }

        /// <summary>
        /// Stretches all selected shapes to the left edge of reference shape
        /// </summary>
        /// <param name="stretchShapes">The shapes to stretch</param>
        public void StretchBottom(PowerPoint.ShapeRange stretchShapes)
        {
            var sa = new StretchAction((PowerPoint.Shape referenceShape, PowerPoint.Shape stretchShape) =>
            {
                // Stretching will cause the object to be very small
                if (stretchShape.Top > GetBottom(referenceShape))
                {
                    return false;
                }
                // The actual stretch action
                stretchShape.Height += GetBottom(referenceShape) - GetBottom(stretchShape);

                return true;
            });
            Stretch(stretchShapes, sa);
        }

        /// <summary>
        /// Stretch shapes in the list
        /// </summary>
        /// <param name="stretchShapes">The shapes to stretch</param>
        /// <param name="stretchAction">The function to use to stretch</param>
        private void Stretch(PowerPoint.ShapeRange stretchShapes, StretchAction stretchAction)
        {
            if (!ValidateSelection(stretchShapes))
            {
                return;
            }

            var referenceShape = GetReferenceShape(stretchShapes);
            var hasStretchedAll = true;

            for (var i = ModShapesIndex; i <= stretchShapes.Count; i++)
            {
                if (!stretchAction(referenceShape, stretchShapes[i]))
                {
                    hasStretchedAll = false;
                }
            }

            ValidateHasStretchedAll(hasStretchedAll);
        }

        #endregion

        #region Helper Functions

        private void ValidateHasStretchedAll(bool hasStretchAll)
        {
            try
            {
                if (!hasStretchAll)
                {
                    ThrowErrorCode(ErrorCodeShapesNotStretchText);
                }
            }
            catch (Exception e)
            {
                ProcessErrorMessage(e);
            }
        }

        private bool ValidateSelection(PowerPoint.ShapeRange shapes)
        {
            if (!IsMoreThanOneShape(shapes))
            {
                return false;
            }

            return true;
        }

        private PowerPoint.Shape GetReferenceShape(PowerPoint.ShapeRange shapes)
        {
            return shapes[1];
        }

        private float GetRight(PowerPoint.Shape aShape)
        {
            return aShape.Left + aShape.Width;
        }

        private float GetBottom(PowerPoint.Shape aShape)
        {
            return aShape.Top + aShape.Height;
        }
        
        #endregion
    }
}
