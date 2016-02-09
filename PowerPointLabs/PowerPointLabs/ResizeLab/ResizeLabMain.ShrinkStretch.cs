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
using PowerPointLabs.Utils;
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
        /// <param name="referenceEdge">The edge to refer to</param>
        /// <param name="stretchShape">The shape to stretch</param>
        /// <returns>True if shape is stretched successfully, false otherwise</returns>
        private delegate void StretchAction(float referenceEdge, PowerPoint.Shape stretchShape);
        
        /// <summary>
        /// Checks whether a shape can be stretched in a particular direction towards reference shape
        /// </summary>
        /// <param name="referenceEdge">The edge to refer to. This may be modified to match the apporiate stretch action</param>
        /// <param name="checkShape">The shape to check</param>
        /// <returns>The appropriate stretch action to use</returns>
        private delegate StretchAction GetAppropriateStretchAction(float referenceEdge, PowerPoint.Shape checkShape);

        /// <summary>
        /// Returns the default reference edge for given shape
        /// </summary>
        /// <param name="referenceShape">The shape to get the reference edge from</param>
        /// <returns>The point determining the reference edge</returns>
        private delegate float GetDefaultReferenceEdge(PowerPoint.Shape referenceShape);
        #region API

        /// <summary>
        /// Stretches all selected shapes to the left edge of reference shape
        /// </summary>
        /// <param name="stretchShapes">The shapes to stretch</param>
        public void StretchLeft(PowerPoint.ShapeRange stretchShapes)
        {
            var appropriateStretch = new GetAppropriateStretchAction((float referenceEdge, PowerPoint.Shape checkShape) =>
            {
                // Opposite stretch
                if (GetRight(checkShape) < referenceEdge)
                {
                    return StretchRightAction;
                }
                return StretchLeftAction;
            });
            var defaultReferenceEdge = new GetDefaultReferenceEdge((PowerPoint.Shape referenceShape) =>
            {
                return Graphics.LeftMostPoint(Graphics.GetRealCoordinates(referenceShape)).X;
            });
            Stretch(stretchShapes, appropriateStretch, defaultReferenceEdge);
        }

        /// <summary>
        /// Stretches all selected shapes to the right edge of reference shape
        /// </summary>
        /// <param name="stretchShapes">The shapes to stretch</param>
        public void StretchRight(PowerPoint.ShapeRange stretchShapes)
        {
            var appropriateStretch = new GetAppropriateStretchAction((float referenceEdge, PowerPoint.Shape checkShape) =>
            {
                // Opposite stretch
                if (checkShape.Left > referenceEdge)
                {
                    return StretchLeftAction;
                }
                return StretchRightAction;
            });
            var defaultReferenceEdge = new GetDefaultReferenceEdge((PowerPoint.Shape referenceShape) =>
            {
                return Graphics.RightMostPoint(Graphics.GetRealCoordinates(referenceShape)).X;
            });
            Stretch(stretchShapes, appropriateStretch, defaultReferenceEdge);
        }

        /// <summary>
        /// Stretches all selected shapes to the top edge of reference shape
        /// </summary>
        /// <param name="stretchShapes">The shapes to stretch</param>
        public void StretchTop(PowerPoint.ShapeRange stretchShapes)
        {
            var appropriateStretch = new GetAppropriateStretchAction((float referenceEdge, PowerPoint.Shape checkShape) =>
            {
                // Opposite stretch
                if (GetBottom(checkShape) < referenceEdge)
                {
                    return StretchBottomAction;
                }
                return StretchTopAction;
            });
            var defaultReferenceEdge = new GetDefaultReferenceEdge((PowerPoint.Shape referenceShape) =>
            {
                return Graphics.TopMostPoint(Graphics.GetRealCoordinates(referenceShape)).Y;
            });
            Stretch(stretchShapes, appropriateStretch, defaultReferenceEdge);
        }

        /// <summary>
        /// Stretches all selected shapes to the left edge of reference shape
        /// </summary>
        /// <param name="stretchShapes">The shapes to stretch</param>
        public void StretchBottom(PowerPoint.ShapeRange stretchShapes)
        {
            var appropriateStretch = new GetAppropriateStretchAction((float referenceEdge, PowerPoint.Shape checkShape) =>
            {
                // Opposite stretch
                if (checkShape.Top > referenceEdge)
                {
                    return StretchTopAction;
                }
                return StretchBottomAction;
            });
            var defaultReferenceEdge = new GetDefaultReferenceEdge((PowerPoint.Shape referenceShape) =>
            {
                return Graphics.BottomMostPoint(Graphics.GetRealCoordinates(referenceShape)).Y;
            });
            Stretch(stretchShapes, appropriateStretch, defaultReferenceEdge);
        }

        private void StretchLeftAction(float referenceEdge, PowerPoint.Shape stretchShape)
        {
            stretchShape.Width += stretchShape.Left - referenceEdge;
            stretchShape.Left = referenceEdge;
        }

        private void StretchRightAction(float referenceEdge, PowerPoint.Shape stretchShape)
        {
            stretchShape.Width += referenceEdge - GetRight(stretchShape);
        }

        private void StretchTopAction(float referenceEdge, PowerPoint.Shape stretchShape)
        {
            stretchShape.Height += stretchShape.Top - referenceEdge;
            stretchShape.Top = referenceEdge;
        }

        private void StretchBottomAction(float referenceEdge, PowerPoint.Shape stretchShape)
        {
            stretchShape.Height += referenceEdge - GetBottom(stretchShape);
        }

        /// <summary>
        /// Stretch shapes in the list
        /// </summary>
        /// <param name="stretchShapes">The shapes to stretch</param>
        /// <param name="stretchAction">The function to use to stretch</param>
        /// <param name="defaultReferenceEdge">The function to return the default reference edge</param>
        private void Stretch(PowerPoint.ShapeRange stretchShapes, GetAppropriateStretchAction stretchAction, GetDefaultReferenceEdge defaultReferenceEdge)
        {
            if (!ValidateSelection(stretchShapes))
            {
                return;
            }

            var referenceShape = GetReferenceShape(stretchShapes);
            float referenceEdge = defaultReferenceEdge(referenceShape);

            for (var i = ModShapesIndex; i <= stretchShapes.Count; i++)
            {
                StretchAction sa = stretchAction(referenceEdge, stretchShapes[i]);
                sa(referenceEdge, stretchShapes[i]);
            }

        }

        #endregion

        #region Helper Functions

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
