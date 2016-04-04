﻿using System;
using Microsoft.Office.Core;
using PowerPointLabs.ActionFramework.Common.Log;
using PowerPointLabs.Utils;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointLabs.ResizeLab
{
    /// <summary>
    /// Match is the partial class of ResizeLabMain.
    /// It handles the resizing of shape's dimension to its respective
    /// width or height
    /// </summary>
    partial class ResizeLabMain
    {
        // To be used for error handling
        internal const int Match_MinNoOfShapesRequired = 1;
        internal const string Match_FeatureName = "Match";
        internal const string Match_ShapeSupport = "object";
        internal static readonly string[] Match_ErrorParameters =
        {
            Match_FeatureName,
            Match_MinNoOfShapesRequired.ToString(),
            Match_ShapeSupport
        };

        /// <summary>
        /// Resize the height of selected shapes to match their width.
        /// </summary>
        /// <param name="selectedShapes"></param>
        public void MatchWidth(PowerPoint.ShapeRange selectedShapes)
        {
            try
            {
                var isAspectRatio = selectedShapes.LockAspectRatio;
                selectedShapes.LockAspectRatio = MsoTriState.msoFalse;

                for (int i = 1; i <= selectedShapes.Count; i++)
                {
                    var shape = new PPShape(selectedShapes[i]);
                    var anchorPoint = GetAnchorPoint(shape);

                    shape.AbsoluteHeight = shape.AbsoluteWidth;
                    AlignShape(shape, anchorPoint);
                }

                selectedShapes.LockAspectRatio = isAspectRatio;
            }
            catch (Exception e)
            {
                Logger.LogException(e, "MatchWidth");
            }
        }

        /// <summary>
        /// Resize the width of selected shapes to match their height.
        /// </summary>
        /// <param name="selectedShapes"></param>
        public void MatchHeight(PowerPoint.ShapeRange selectedShapes)
        {
            try
            {
                var isAspectRatio = selectedShapes.LockAspectRatio;
                selectedShapes.LockAspectRatio = MsoTriState.msoFalse;

                for (int i = 1; i <= selectedShapes.Count; i++)
                {
                    var shape = new PPShape(selectedShapes[i]);
                    var anchorPoint = GetAnchorPoint(shape);

                    shape.AbsoluteWidth = shape.AbsoluteHeight;
                    AlignShape(shape, anchorPoint);
                }

                selectedShapes.LockAspectRatio = isAspectRatio;
            }
            catch (Exception e)
            {
                Logger.LogException(e, "MatchHeight");
            }
        }
    }
}
