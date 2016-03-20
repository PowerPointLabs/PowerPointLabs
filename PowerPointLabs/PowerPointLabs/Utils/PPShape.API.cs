using System;
using Microsoft.Office.Core;
using System.Drawing;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointLabs.Utils
{
    public partial class PPShape
    {
        private readonly PowerPoint.Shape _shape;
        private float _absoluteWidth;
        private float _absoluteHeight;
        private float _rotatedLeft;
        private float _rotatedTop;
        private float _originalRotation;

        public PPShape(PowerPoint.Shape shape)
        {
            _shape = shape;
            _originalRotation = _shape.Rotation;

            ConvertToFreeform();

            UpdateAbsoluteWidth();
            UpdateAbsoluteHeight();

            UpdateTop();
            UpdateLeft();
        }

        #region Properties

        /// <summary>
        /// Return or set the name of the specified shape.
        /// </summary>
        public string Name
        {
            get { return _shape.Name; }
            set { _shape.Name = value; }
        }

        /// <summary>
        /// Return a 64-bit signed integer that identifies the PPshape. Read-only.
        /// </summary>
        public int Id => _shape.Id;

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
                
                if (_shape.LockAspectRatio == MsoTriState.msoTrue)
                {
                    SetToAbsoluteWidthAspectRatio();
                }
                else
                {
                    SetToAbsoluteDimension();
                }     
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

                if (_shape.LockAspectRatio == MsoTriState.msoTrue)
                {
                    SetToAbsoluteHeightAspectRatio();
                }
                else
                {
                    SetToAbsoluteDimension();
                }
            }
        }

        /// <summary>
        /// Return or set the shape type for the specified Shape object,
        /// which must represent an AutoShape other than a line, freeform drawing, or connector.
        /// Read/write.
        /// </summary>
        public MsoAutoShapeType AutoShapeType
        {
            get { return _shape.AutoShapeType; }
            set { _shape.AutoShapeType = value; }
        }

        /// <summary>
        /// Return a point that represents the center of the shape.
        /// </summary>
        public PointF Center
        {
            get
            {
                var centerPoint = new PointF
                {
                    X = _rotatedLeft + _absoluteWidth/2,
                    Y = _rotatedTop + _absoluteHeight/2
                };
                return centerPoint;
            }
        }

        /// <summary>
        /// Return a point that represents the top left of the shape's bounding box after rotation.
        /// </summary>
        public PointF TopLeft
        {
            get
            {
                var topLeft = new PointF
                {
                    X = _rotatedLeft,
                    Y = _rotatedTop
                };
                return topLeft;
            }
        }

        /// <summary>
        /// Return a point that represents the top center of the shape's bounding box after rotation.
        /// </summary>
        public PointF TopCenter
        {
            get
            {
                var topCenterPoint = new PointF
                {
                    X = _rotatedLeft + _absoluteWidth / 2,
                    Y = _rotatedTop
                };
                return topCenterPoint;
            }
        }

        /// <summary>
        /// Return a point that represents the top right of the shape's bounding box after rotation.
        /// </summary>
        public PointF TopRight
        {
            get
            {
                var topRightPoint = new PointF
                {
                    X = _rotatedLeft + _absoluteWidth,
                    Y = _rotatedTop
                };
                return topRightPoint;
            }
        }

        /// <summary>
        /// Return a point that represents the middle left of the shape's bounding box after rotation.
        /// </summary>
        public PointF MiddleLeft
        {
            get
            {
                var middleLeftPoint = new PointF
                {
                    X = _rotatedLeft,
                    Y = _rotatedTop + _absoluteHeight / 2
                };
                return middleLeftPoint;
            }
        }

        /// <summary>
        /// Return a point that represents the middle right of the shape's bounding box after rotation.
        /// </summary>
        public PointF MiddleRight
        {
            get
            {
                var middleRightPoint = new PointF
                {
                    X = _rotatedLeft + _absoluteWidth,
                    Y = _rotatedTop + _absoluteHeight / 2
                };
                return middleRightPoint;
            }
        }

        /// <summary>
        /// Return a point that represents the bottom left of the shape's bounding box after rotation.
        /// </summary>
        public PointF BottomLeft
        {
            get
            {
                var bottomLeftPoint = new PointF
                {
                    X = _rotatedLeft,
                    Y = _rotatedTop + _absoluteHeight
                };
                return bottomLeftPoint;
            }
        }

        /// <summary>
        /// Return a point that represents the bottom center of the shape's bounding box after rotation.
        /// </summary>
        public PointF BottomCenter
        {
            get
            {
                var bottomCenterPoint = new PointF
                {
                    X = _rotatedLeft + _absoluteWidth / 2,
                    Y = _rotatedTop + _absoluteHeight
                };
                return bottomCenterPoint;
            }
        }

        /// <summary>
        /// Return a point that represents the bottom right of the shape's bounding box after rotation.
        /// </summary>
        public PointF BottomRight
        {
            get
            {
                var bottomRightPoint = new PointF
                {
                    X = _rotatedLeft + _absoluteWidth,
                    Y = _rotatedTop + _absoluteHeight
                };
                return bottomRightPoint;
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
        /// Return or set the degrees of specified shape is rotated around the z-axis. 
        /// Read/write.
        /// </summary>
        public float ShapeRotation
        {
            get { return _originalRotation; }
            set { _originalRotation = value; }
        }

        /// <summary>
        /// Return or set the degrees of specified shape's bounding box is rotated around the z-axis. 
        /// Read/write.
        /// </summary>
        public float BoxRotation
        {
            get { return _shape.Rotation; }
            set
            {
                _shape.Rotation = value;
                ConvertToFreeform();
                UpdateAbsoluteHeight();
                UpdateAbsoluteWidth();
                UpdateLeft();
                UpdateTop();
            }
        }

        /// <summary>
        /// Returns the position of the specified shape in the z-order. Read-only.
        /// </summary>
        public int ZOrderPosition => _shape.ZOrderPosition;

        #endregion

        #region Functions

        /// <summary>
        /// Delete the specified Shape object.
        /// </summary>
        public void Delete()
        {
            _shape.Delete();
        }

        /// <summary>
        /// Create a duplicate of the specified Shape object and return a new shape.
        /// </summary>
        /// <returns></returns>
        public PPShape Duplicate()
        {
            var newShape = new PPShape(_shape.Duplicate()[1]) {Name = _shape.Name + "Copy"};
            return newShape;
        }

        /// <summary>
        /// Moves the specified shape horizontally by the specified number of points.
        /// </summary>
        /// <param name="value">Number of points from left of slide</param>
        public void IncrementLeft(float value)
        {
            _shape.IncrementLeft(value);
            UpdateLeft();
        }

        /// <summary>
        /// Moves the specified shape vertically by the specified number of points.
        /// </summary>
        /// <param name="value">Number of points from top of slide</param>
        public void IncrementTop(float value)
        {
            _shape.IncrementTop(value);
            UpdateTop();
        }

        /// <summary>
        /// Flip the specified shape around its horizontal or vertical axis.
        /// </summary>
        /// <param name="msoFlipCmd"></param>
        public void Flip(MsoFlipCmd msoFlipCmd)
        {
            _shape.Flip(msoFlipCmd);
        }

        /// <summary>
        /// Select the specified object.
        /// </summary>
        /// <param name="replace"></param>
        public void Select(MsoTriState replace)
        {
            _shape.Select(replace);
        }

        /// <summary>
        /// Reset the nodes to corresponding original rotation.
        /// </summary>
        public void ResetNodes()
        {
            if (_shape.Type != MsoShapeType.msoFreeform || _shape.Nodes.Count < 1) return;

            var rotation = GetStandardizedRotation(360 - _originalRotation);
            var centerLeft = Center.X;
            var centerTop = Center.Y;

            for (int i = 1; i <= _shape.Nodes.Count; i++)
            {
                var node = _shape.Nodes[i];
                var point = node.Points;
                var oldX = point[1, 1];
                var oldY = point[1, 2];
                var newX = oldY*Math.Sin(rotation) + oldX*Math.Cos(rotation);
                var newY = oldY*Math.Cos(rotation) - oldX*Math.Sin(rotation);

                _shape.Nodes.SetPosition(i, newX, newY);
            }

            _shape.Rotation = _originalRotation;

            UpdateAbsoluteWidth();
            UpdateAbsoluteHeight();

            Left = centerLeft - _absoluteWidth/2;
            Top = centerTop - _absoluteHeight/2;
        }

        #endregion
    }
}
