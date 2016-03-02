using System;
using Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointLabs.Utils
{
    internal class PPShape
    {
        private PowerPoint.Shape _shape;
        private float _absoluteWidth;
        private float _absoluteHeight;
        private float _rotatedLeft;
        private float _rotatedTop;

        public PPShape(PowerPoint.Shape shape)
        {
            _shape = shape;

            UpdateAbsoluteWidth();
            UpdateAbsoluteHeight();

            UpdateTop();
            UpdateLeft();
        }

        /// <summary>
        /// Return or set the width of the specified shape.
        /// </summary>
        public float ShapeWidth
        {
            get { return _shape.Width; }
            set
            {
                _shape.Width = value;
                UpdateAbsoluteWidth();
            }
        }

        /// <summary>
        /// Return or set the height of the specified shape.
        /// </summary>
        public float ShapeHeight
        {
            get { return _shape.Height; }
            set
            {
                _shape.Height = value;
                UpdateAbsoluteHeight();
            }
        }

        /// <summary>
        /// Return or set the absolute width of rotated shape.
        /// </summary>
        public float AbsoluteWidth
        {
            get { return _absoluteWidth; }
            set
            {
                _absoluteWidth = value;
                SetToAbsoluteDimension();
            }
        }

        /// <summary>
        /// Return or set the absolute height of rotated shape.
        /// </summary>
        public float AbsoluteHeight
        {
            get { return _absoluteHeight; }
            set
            {
                _absoluteHeight = value;
                SetToAbsoluteDimension();
            }
        }

        /// <summary>
        /// Return or set a single-precision floating-point number that represents the 
        /// distance from the top most point of the shape to the top edge of the slide.
        /// </summary>
        public float Top
        {
            get { return _rotatedTop; }
            set
            {
                _rotatedTop = value; 
                SetTop();
            }
        }


        /// <summary>
        /// Return or set a single-precision floating-point number that represents the 
        /// distance from the left most point of the shape to the left edge of the slide.
        /// </summary>
        public float Left
        {
            get { return _rotatedLeft; }
            set
            {
                _rotatedLeft = value; 
                SetLeft();
            }
        }

        // Only work for sharp edge
        private void ConvertToFreeform(PowerPoint.Slide slide)
        {
            if (_shape.Type != MsoShapeType.msoAutoShape && _shape.Nodes.Count > 0) return;
            _shape.Nodes.Insert(1, MsoSegmentType.msoSegmentLine, MsoEditingType.msoEditingAuto, 0, 0);
            _shape.Nodes.Delete(2);

            var nodes = _shape.Nodes;
            var firstNode = nodes[1];
            var firstPoint = firstNode.Points;
 
            PowerPoint.FreeformBuilder freeform = slide.Shapes.BuildFreeform(firstNode.EditingType, firstPoint[1, 1], firstPoint[1, 2]);


            for (int i = 2; i <= nodes.Count; i++)
            {
                var node = nodes[i];
                var point = node.Points;

                try
                {
                    freeform.AddNodes(node.SegmentType, node.EditingType, point[1, 1], point[1, 2]);
                }
                catch (Exception)
                {
                    freeform.AddNodes(node.SegmentType, MsoEditingType.msoEditingAuto, point[1, 1], point[1, 2]);
                }

            }

            _shape.Delete();
            _shape = freeform.ConvertToShape();
        }
    

        /// <summary>
        /// Update the absolute width according to the actual shape width and height.
        /// </summary>
        private void UpdateAbsoluteWidth()
        {
            var rotation = GetStandardizedRotation(_shape.Rotation);

            if (IsInQuadrant(_shape.Rotation))
            {
                _absoluteWidth = (float) (_shape.Height*Math.Sin(rotation) + _shape.Width*Math.Cos(rotation));
            }
            else if ((int) _shape.Rotation == 90 || (int) _shape.Rotation == 270)
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
                _absoluteHeight = (float) (_shape.Height*Math.Cos(rotation) + _shape.Width*Math.Sin(rotation));
            }
            else if ((int) _shape.Rotation == 90 || (int) _shape.Rotation == 270)
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
            _rotatedTop = _shape.Top + _shape.Height/2 - _absoluteHeight/2;
        }

        /// <summary>
        /// Update the distance from left most point of the shape to left edge of the slide.
        /// </summary>
        private void UpdateLeft()
        {
            _rotatedLeft = _shape.Left + _shape.Width/2 - _absoluteWidth/2;
        }

        /// <summary>
        /// Set the actual width and height according to the absolute dimension (e.g. width and height).
        /// </summary>
        private void SetToAbsoluteDimension()
        {
            var rotation = GetStandardizedRotation(_shape.Rotation);
            var sinAngle = Math.Sin(rotation);
            var cosAngle = Math.Cos(rotation);
            var ratio = sinAngle/cosAngle;

            _shape.Height = (float) ((_absoluteWidth*ratio - _absoluteHeight)/(sinAngle*ratio - cosAngle));
            _shape.Width = (float) ((_absoluteWidth - _shape.Height*sinAngle)/cosAngle);
        }

        /// <summary>
        /// Set the distance from the top edge of unrotated shape to the top edge of the slide.
        /// </summary>
        private void SetTop()
        {
            _shape.Top = _rotatedTop - _shape.Height/2 + _absoluteHeight/2;
        }

        /// <summary>
        /// Set the distance from the left edge of unrotated shape to the left edge of the slide.
        /// </summary>
        private void SetLeft()
        {
            _shape.Left = _rotatedLeft - _shape.Width/2 + _absoluteWidth/2;
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
                rotation = rotation%90;
            }
            else if ((rotation > 90 && rotation <= 180) ||
                     (rotation > 270 && rotation <= 360))
            {
                rotation = (360 - rotation)%90;
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
            return (float) (rotation*Math.PI/180);
        }
    }
}
