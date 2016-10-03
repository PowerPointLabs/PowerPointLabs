using System;
using System.Collections.Generic;
using System.Drawing;
using Microsoft.Office.Core;

namespace PowerPointLabs.Utils
{
    partial class PPShape
    {
        /// <summary>
        /// Convert Autoshape to freeform
        /// </summary>
        private void ConvertToFreeform()
        {
            if ((int)_shape.Rotation == 0) return;
            
            SetPoints(isConvertToFreeform: true);
            if (_points == null)
            {
                return;
            }

            // Rotate bounding box back to 0 degree, 
            // flip the shape to original orientation 
            // and apply the original coordinates to the nodes
            _shape.Rotation = 0;

            if (_shape.VerticalFlip == MsoTriState.msoTrue)
            {
                _shape.Flip(MsoFlipCmd.msoFlipVertical);
            }

            if (_shape.HorizontalFlip == MsoTriState.msoTrue)
            {
                _shape.Flip(MsoFlipCmd.msoFlipHorizontal);
            }

            for (int i = 0; i < _points.Count; i++)
            {
                var point = _points[i];
                var nodeIndex = i + 1;

                _shape.Nodes.SetPosition(nodeIndex, point.X, point.Y);
            }
        }

        /// <summary>
        /// Update the absolute width according to the actual shape width and height.
        /// </summary>
        private void UpdateAbsoluteWidth()
        {
            var rotation = GetStandardizedRotation(_shape.Rotation);

            if (IsInQuadrant(_shape.Rotation))
            {
                _absoluteWidth = (float)(_shape.Height * Math.Sin(rotation) + _shape.Width * Math.Cos(rotation));
            }
            else if ((int)_shape.Rotation == 90 || (int)_shape.Rotation == 270)
            {
                _absoluteWidth = _shape.Height;
            }
            else
            {
                _absoluteWidth = _shape.Width;
            }
        }

        /// <summary>
        /// Update the absolute height according to the actual shape width and height.
        /// </summary>
        private void UpdateAbsoluteHeight()
        {
            var rotation = GetStandardizedRotation(_shape.Rotation);

            if (IsInQuadrant(_shape.Rotation))
            {
                _absoluteHeight = (float)(_shape.Height * Math.Cos(rotation) + _shape.Width * Math.Sin(rotation));
            }
            else if ((int)_shape.Rotation == 90 || (int)_shape.Rotation == 270)
            {
                _absoluteHeight = _shape.Width;
            }
            else
            {
                _absoluteHeight = _shape.Height;
            }
        }

        /// <summary>
        /// Update the distance from top most point of the shape to top edge of the slide.
        /// </summary>
        private void UpdateVisualTop()
        {
            _rotatedVisualTop = _shape.Top + _shape.Height / 2 - _absoluteHeight / 2;
        }

        /// <summary>
        /// Update the distance from left most point of the shape to left edge of the slide.
        /// </summary>
        private void UpdateVisualLeft()
        {
            _rotatedVisualLeft = _shape.Left + _shape.Width / 2 - _absoluteWidth / 2;
        }

        /// <summary>
        /// Set the actual width and height according to the absolute dimension (e.g. width and height).
        /// </summary>
        private void SetToAbsoluteDimension()
        {
            var rotation = GetStandardizedRotation(_shape.Rotation);
            var sinAngle = Math.Sin(rotation);
            var cosAngle = Math.Cos(rotation);

            if ((int) _shape.Rotation == 90 || (int) _shape.Rotation == 270)
            {
                _shape.Height = _absoluteWidth;
                _shape.Width = _absoluteHeight;
            }
            else
            {
                var ratio = sinAngle / cosAngle;
                _shape.Height = (float)((_absoluteWidth * ratio - _absoluteHeight) / (sinAngle * ratio - cosAngle));
                _shape.Width = (float)((_absoluteWidth - _shape.Height * sinAngle) / cosAngle);
            }
        }

        private void SetToAbsoluteHeightAspectRatio()
        {
            // Store the original position of the shape
            var originalTop = _shape.Top;
            var originalLeft = _shape.Left;

            _shape.LockAspectRatio = MsoTriState.msoFalse;
            FitToSlide.FitToHeight(_shape, _absoluteWidth, _absoluteHeight);
            _shape.LockAspectRatio = MsoTriState.msoTrue;

            _shape.Top = originalTop;
            _shape.Left = originalLeft;

            UpdateAbsoluteWidth();
            UpdateAbsoluteHeight();
        }

        private void SetToAbsoluteWidthAspectRatio()
        {
            // Store the original position of the shape
            var originalTop = _shape.Top;
            var originalLeft = _shape.Left;

            _shape.LockAspectRatio = MsoTriState.msoFalse;
            FitToSlide.FitToWidth(_shape, _absoluteWidth, _absoluteHeight);
            _shape.LockAspectRatio = MsoTriState.msoTrue;

            _shape.Top = originalTop;
            _shape.Left = originalLeft;

            UpdateAbsoluteWidth();
            UpdateAbsoluteHeight();
        }

        /// <summary>
        /// Set the distance from the top edge of unrotated shape to the top edge of the slide.
        /// </summary>
        private void SetTop()
        {
            _shape.Top = _rotatedVisualTop - _shape.Height / 2 + _absoluteHeight / 2;
        }

        /// <summary>
        /// Set the distance from the left edge of unrotated shape to the left edge of the slide.
        /// </summary>
        private void SetLeft()
        {
            _shape.Left = _rotatedVisualLeft - _shape.Width / 2 + _absoluteWidth / 2;
        }

        /// <summary>
        /// Save the coordinates of nodes
        /// </summary>
        private void SetPoints(bool isConvertToFreeform = false)
        {
            if (!(_shape.Type == MsoShapeType.msoAutoShape || _shape.Type == MsoShapeType.msoFreeform)
                || _shape.Nodes.Count < 1)
            {
                return;
            }

            var shape = _shape;

            if (!isConvertToFreeform)
            {
                shape = _shape.Duplicate()[1];
                shape.Left = _shape.Left;
                shape.Top = _shape.Top;
            }

            // Convert AutoShape to Freeform shape
            if (shape.Type == MsoShapeType.msoAutoShape)
            {
                shape.Nodes.Insert(1, MsoSegmentType.msoSegmentLine, MsoEditingType.msoEditingAuto, 0, 0);
                shape.Nodes.Delete(2);
            }

            _points = new List<PointF>();

            for (int i = 1; i <= shape.Nodes.Count; i++)
            {
                var node = shape.Nodes[i];
                var point = node.Points;
                var newPoint = new PointF(point[1, 1], point[1, 2]);

                _points.Add(newPoint);
            }

            if (!isConvertToFreeform)
            {
                shape.Delete();
            }
        }

        /// <summary>
        /// Save the coordinates of the bounding box nodes
        /// </summary>
        private void SetBoundingBoxPoints()
        {
            _points = new List<PointF>();

            _points.Add(VisualTopLeft);
            _points.Add(VisualTopCenter);
            _points.Add(VisualTopRight);
            _points.Add(VisualMiddleRight);
            _points.Add(VisualBottomRight);
            _points.Add(VisualBottomCenter);
            _points.Add(VisualBottomLeft);
            _points.Add(VisualMiddleLeft);
            _points.Add(VisualTopLeft);
        }

        /// <summary>
        /// Check if the angle is in the quadrant.
        /// </summary>
        /// <param name="rotation"></param>
        /// <returns></returns>
        private static bool IsInQuadrant(float rotation)
        {
            return (rotation > 0 && rotation < 90) || (rotation > 90 && rotation < 180) ||
                   (rotation > 180 && rotation < 270) || (rotation > 270 && rotation < 360);
        }

        /// <summary>
        /// Standardize the angle to the first quadrant.
        /// </summary>
        /// <param name="rotation"></param>
        /// <returns></returns>
        private static float GetStandardizedRotation(float rotation)
        {
            if ((rotation > 0 && rotation < 90) ||
                (rotation > 180 && rotation < 270))
            {
                rotation = rotation % 90;
            }
            else if ((rotation > 90 && rotation <= 180) ||
                     (rotation > 270 && rotation <= 360))
            {
                rotation = (360 - rotation) % 90;
            }
            else if ((int)rotation == 270)
            {
                rotation = 360 - rotation;
            }

            return ConvertDegToRad(rotation);
        }

        /// <summary>
        /// Convert angle from degree to radian.
        /// </summary>
        /// <param name="rotation"></param>
        /// <returns></returns>
        private static float ConvertDegToRad(float rotation)
        {
            return (float)(rotation * Math.PI / 180);
        }

        /// <summary>
        /// Get the point after rotation with the reference to center.
        /// </summary>
        /// <param name="widthDiff"></param>
        /// <param name="heightDiff"></param>
        /// <returns></returns>
        private PointF GetRotatedPoint(float widthDiff, float heightDiff)
        {
            var rotation = ConvertDegToRad(_shape.Rotation);
            var centerX = _shape.Left + _shape.Width/2;
            var centerY = _shape.Top + _shape.Height/2;
            var x = Math.Cos(rotation)*widthDiff - Math.Sin(rotation)*heightDiff + centerX;
            var y = Math.Sin(rotation)*widthDiff + Math.Cos(rotation)*heightDiff + centerY;

            return new PointF((float) x, (float) y);
        }
        

        /// <summary>
        /// Get the center of the shape.
        /// </summary>
        /// <param name="rotated"></param>
        /// <param name="widthDiff"></param>
        /// <param name="heightDiff"></param>
        /// <returns></returns>
        private PointF GetCenterPoint(PointF rotated, float widthDiff, float heightDiff)
        {
            var rotation = ConvertDegToRad(_shape.Rotation);
            var x = rotated.X - Math.Cos(rotation)*widthDiff + Math.Sin(rotation)*heightDiff;
            var y = rotated.Y - Math.Sin(rotation)*widthDiff - Math.Cos(rotation)*heightDiff;
            
            return new PointF((float) x, (float) y);
        }

        /// <summary>
        /// Align the shape to position with regards to the center.
        /// </summary>
        /// <param name="center"></param>
        private void AlignToCenter(PointF center)
        {
            _shape.Left = center.X - _shape.Width/2;
            _shape.Top = center.Y - _shape.Height/2;
        }
    }
}
