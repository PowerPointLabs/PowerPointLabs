using System;
using System.Collections;
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
            if (!(_shape.Type == MsoShapeType.msoAutoShape || _shape.Type == MsoShapeType.msoFreeform) || _shape.Nodes.Count < 1) return;

            // Convert AutoShape to Freeform shape
            if (_shape.Type == MsoShapeType.msoAutoShape)
            {
                _shape.Nodes.Insert(1, MsoSegmentType.msoSegmentLine, MsoEditingType.msoEditingAuto, 0, 0);
                _shape.Nodes.Delete(2);
            }

            // Save the coordinates of nodes
            var pointList = new ArrayList();
            for (int i = 1; i <= _shape.Nodes.Count; i++)
            {
                var node = _shape.Nodes[i];
                var point = node.Points;
                var newPoint = new float[] { point[1, 1], point[1, 2] };

                pointList.Add(newPoint);
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

            for (int i = 0; i < pointList.Count; i++)
            {
                var point = (float[])pointList[i];
                var nodeIndex = i + 1;

                _shape.Nodes.SetPosition(nodeIndex, point[0], point[1]);
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
        private void UpdateTop()
        {
            _rotatedTop = _shape.Top + _shape.Height / 2 - _absoluteHeight / 2;
        }

        /// <summary>
        /// Update the distance from left most point of the shape to left edge of the slide.
        /// </summary>
        private void UpdateLeft()
        {
            _rotatedLeft = _shape.Left + _shape.Width / 2 - _absoluteWidth / 2;
        }

        /// <summary>
        /// Set the actual width and height according to the absolute dimension (e.g. width and height).
        /// </summary>
        private void SetToAbsoluteDimension()
        {
            var rotation = GetStandardizedRotation(_shape.Rotation);
            var sinAngle = Math.Sin(rotation);
            var cosAngle = Math.Cos(rotation);

            if ((int) cosAngle != 0)
            {
                var ratio = sinAngle/cosAngle;

                _shape.Height = (float) ((_absoluteWidth*ratio - _absoluteHeight)/(sinAngle*ratio - cosAngle));
                _shape.Width = (float) ((_absoluteWidth - _shape.Height*sinAngle)/cosAngle);
            }
            else
            {
                _shape.Height = _absoluteWidth;
                _shape.Width = _absoluteHeight;
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
            _shape.Top = _rotatedTop - _shape.Height / 2 + _absoluteHeight / 2;
        }

        /// <summary>
        /// Set the distance from the left edge of unrotated shape to the left edge of the slide.
        /// </summary>
        private void SetLeft()
        {
            _shape.Left = _rotatedLeft - _shape.Width / 2 + _absoluteWidth / 2;
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
    }
}
