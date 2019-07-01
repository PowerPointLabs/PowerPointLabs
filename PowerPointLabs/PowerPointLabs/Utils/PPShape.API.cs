using System;
using System.Collections.Generic;
using System.Drawing;

using Microsoft.Office.Core;
using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointLabs.Utils
{
    public partial class PPShape
    {
        public readonly PowerPoint.Shape _shape;
        private float _absoluteWidth;
        private float _absoluteHeight;
        private float _rotatedVisualLeft;
        private float _rotatedVisualTop;
        private float _originalRotation;
        private List<PointF> _points;

        public PPShape(PowerPoint.Shape shape, bool redefineBoundingBox = true)
        {
            _shape = shape;
            _originalRotation = _shape.Rotation;

            if (redefineBoundingBox && (int) _shape.Rotation%90 != 0)
            {
                ConvertToFreeform();
            }
            else
            {
                SetPoints();
            }

            UpdateAbsoluteWidth();
            UpdateAbsoluteHeight();

            UpdateVisualTop();
            UpdateVisualLeft();

            if (_points == null)
            {
                SetBoundingBoxPoints();
            }
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
        /// Return or set a point that represents the visual center of the shape.
        /// </summary>
        public PointF VisualCenter
        {
            get { return new PointF(_rotatedVisualLeft + _absoluteWidth/2, _rotatedVisualTop + _absoluteHeight/2); }
            set
            {
                VisualLeft = value.X - AbsoluteWidth/2;
                VisualTop = value.Y - AbsoluteHeight/2;
            }
        }

        /// <summary>
        /// Return or set a point that represents the actual center of the shape.
        /// </summary>
        public PointF ActualCenter
        {
            get { return GetRotatedPoint(0, 0); }
            set { AlignToCenter(value); }
        }

        /// <summary>
        /// Return or set a point that represents the visual top left of the shape's bounding box after rotation.
        /// </summary>
        public PointF VisualTopLeft
        {
            get { return new PointF(_rotatedVisualLeft, _rotatedVisualTop); }
            set
            {
                VisualLeft = value.X;
                VisualTop = value.Y;
            }
        }

        /// <summary>
        /// Return or set a point that represents the actual top left of the shape's bounding box after rotation.
        /// </summary>
        public PointF ActualTopLeft
        {
            get { return GetRotatedPoint(-_shape.Width/2, -_shape.Height/2); }
            set
            {
                PointF center = GetCenterPoint(value, -_shape.Width/2, -_shape.Height/2);
                AlignToCenter(center);
            }
        }

        /// <summary>
        /// Return or set a point that represents the visual top center of the shape's bounding box after rotation.
        /// </summary>
        public PointF VisualTopCenter
        {
            get { return new PointF(_rotatedVisualLeft + _absoluteWidth/2, _rotatedVisualTop); }
            set
            {
                VisualLeft = value.X - AbsoluteWidth/2;
                VisualTop = value.Y;
            }
        }

        /// <summary>
        /// Return or set a point that represents the actual top center of the shape's bounding box after rotation.
        /// </summary>
        public PointF ActualTopCenter
        {
            get { return GetRotatedPoint(0, -_shape.Height/2); }
            set
            {
                PointF center = GetCenterPoint(value, 0, -_shape.Height/2);
                AlignToCenter(center);
            }
        }

        /// <summary>
        /// Return or set a point that represents the visual top right of the shape's bounding box after rotation.
        /// </summary>
        public PointF VisualTopRight
        {
            get { return new PointF(_rotatedVisualLeft + _absoluteWidth, _rotatedVisualTop); }
            set
            {
                VisualLeft = value.X - AbsoluteWidth;
                VisualTop = value.Y;
            }
        }

        /// <summary>
        /// Return or set a point that represents the actual top right of the shape's bounding box after rotation.
        /// </summary>
        public PointF ActualTopRight
        {
            get
            {
                return GetRotatedPoint(_shape.Width/2, -_shape.Height/2);
            }
            set
            {
                PointF center = GetCenterPoint(value, _shape.Width/2, -_shape.Height/2);
                AlignToCenter(center);
            }
        }

        /// <summary>
        /// Return or set a point that represents the visual middle left of the shape's bounding box after rotation.
        /// </summary>
        public PointF VisualMiddleLeft
        {
            get { return new PointF(_rotatedVisualLeft, _rotatedVisualTop + _absoluteHeight / 2); }
            set
            {
                VisualLeft = value.X;
                VisualTop = value.Y - AbsoluteHeight/2;
            }
        }

        /// <summary>
        /// Retur or setn a point that represents the actual middle left of the shape's bounding box after rotation.
        /// </summary>
        public PointF ActualMiddleLeft
        {
            get { return GetRotatedPoint(-_shape.Width/2, 0); }
            set
            {
                PointF center = GetCenterPoint(value, -_shape.Width/2, 0); 
                AlignToCenter(center);
            }
        }

        /// <summary>
        /// Return or set a point that represents the visual middle right of the shape's bounding box after rotation.
        /// </summary>
        public PointF VisualMiddleRight
        {
            get { return new PointF(_rotatedVisualLeft + _absoluteWidth, _rotatedVisualTop + _absoluteHeight/2); }
            set
            {
                VisualLeft = value.X - AbsoluteWidth;
                VisualTop = value.Y - AbsoluteHeight/2;
            }
        }

        /// <summary>
        /// Return or set a point that represents the actual middle right of the shape's bounding box after rotation.
        /// </summary>
        public PointF ActualMiddleRight
        {
            get { return GetRotatedPoint(_shape.Width/2, 0); }
            set
            {
                PointF center = GetCenterPoint(value, _shape.Width/2, 0);
                AlignToCenter(center);
            }
        }

        /// <summary>
        /// Return or set a point that represents the visual bottom left of the shape's bounding box after rotation.
        /// </summary>
        public PointF VisualBottomLeft
        {
            get { return new PointF(_rotatedVisualLeft, _rotatedVisualTop + _absoluteHeight); }
            set
            {
                VisualLeft = value.X;
                VisualTop = value.Y - AbsoluteHeight;
            }
        }

        /// <summary>
        /// Return or set a point that represents the actual bottom left of the shape's bounding box after rotation.
        /// </summary>
        public PointF ActualBottomLeft
        {
            get { return GetRotatedPoint(-_shape.Width/2, _shape.Height/2); }
            set
            {
                PointF center = GetCenterPoint(value, -_shape.Width/2, _shape.Height/2);
                AlignToCenter(center);
            }
        }

        /// <summary>
        /// Return or set a point that represents the visual bottom center of the shape's bounding box after rotation.
        /// </summary>
        public PointF VisualBottomCenter
        {
            get { return new PointF(_rotatedVisualLeft + _absoluteWidth/2, _rotatedVisualTop + _absoluteHeight); }
            set
            {
                VisualLeft = value.X - AbsoluteWidth/2;
                VisualTop = value.Y - AbsoluteHeight;
            }
        }

        /// <summary>
        /// Return or set a point that represents the actual bottom center of the shape's bounding box after rotation.
        /// </summary>
        public PointF ActualBottomCenter
        {
            get { return GetRotatedPoint(0, _shape.Height/2); }
            set
            {
                PointF center = GetCenterPoint(value, 0, _shape.Height/2);
                AlignToCenter(center);
            }
        }

        /// <summary>
        /// Return or set a point that represents the visual bottom right of the shape's bounding box after rotation.
        /// </summary>
        public PointF VisualBottomRight
        {
            get { return new PointF(_rotatedVisualLeft + _absoluteWidth, _rotatedVisualTop + _absoluteHeight); }
            set
            {
                VisualLeft = value.X - AbsoluteWidth;
                VisualTop = value.Y - AbsoluteHeight;
            }
        }

        /// <summary>
        /// Return or set a point that represents the actual bottom right of the shape's bounding box after rotation.
        /// </summary>
        public PointF ActualBottomRight
        {
            get { return GetRotatedPoint(_shape.Width/2, _shape.Height/2); }
            set
            {
                PointF center = GetCenterPoint(value, _shape.Width/2, _shape.Height/2);
                AlignToCenter(center);
            }
        }

        /// <summary>
        /// Return or set a single-precision floating-point number that represents the 
        /// distance from the left most point of the shape to the left edge of the slide.
        /// </summary>
        public float VisualLeft
        {
            get { return _rotatedVisualLeft; }
            set
            {
                _rotatedVisualLeft = value; 
                SetLeft();
            }
        }

        /// <summary>
        /// Return or set a single-precision floating-point number that represents the 
        /// distance from the top most point of the shape to the top edge of the slide.
        /// </summary>
        public float VisualTop
        {
            get { return _rotatedVisualTop; }
            set
            {
                _rotatedVisualTop = value;
                SetTop();
            }
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
                UpdateVisualLeft();
                UpdateVisualTop();
            }
        }

        /// <summary>
        /// Returns the coordinates of nodes.
        /// </summary>
        public List<PointF> Points => _points;

        /// <summary>
        /// Return or set the position of the specified shape in the z-order
        /// Read/write
        /// </summary>
        public int ZOrderPosition
        {
            get { return _shape.ZOrderPosition; }
        }

        #endregion

        #region Functions

        /// <summary>
        /// Delete the specified Shape object.
        /// </summary>
        public void Delete()
        {
            _shape.SafeDelete();
        }

        /// <summary>
        /// Create a duplicate of the specified Shape object and return a new shape.
        /// </summary>
        /// <returns></returns>
        public PPShape Duplicate()
        {
            PPShape newShape = new PPShape(_shape.Duplicate()[1]) {Name = _shape.Name + "Copy"};
            return newShape;
        }

        /// <summary>
        /// Moves the specified shape horizontally by the specified number of points.
        /// </summary>
        /// <param name="value">Number of points from left of slide</param>
        public void IncrementLeft(float value)
        {
            _shape.IncrementLeft(value);
            UpdateVisualLeft();
        }

        /// <summary>
        /// Moves the specified shape vertically by the specified number of points.
        /// </summary>
        /// <param name="value">Number of points from top of slide</param>
        public void IncrementTop(float value)
        {
            _shape.IncrementTop(value);
            UpdateVisualTop();
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
            if (_shape.Type != MsoShapeType.msoFreeform || _shape.Nodes.Count < 1)
            {
                return;
            }

            bool isSecondOrFourthQuadrant = (_originalRotation >= 90 && _originalRotation < 180) ||
                                         (_originalRotation >= 270 && _originalRotation < 360);

            float rotation = GetStandardizedRotation(_originalRotation%90);
            float centerLeft = VisualCenter.X;
            float centerTop = VisualCenter.Y;

            for (int i = 1; i <= _shape.Nodes.Count; i++)
            {
                PowerPoint.ShapeNode node = _shape.Nodes[i];
                dynamic point = node.Points;
                dynamic oldX = point[1, 1];
                dynamic oldY = point[1, 2];
                dynamic newX = oldY*Math.Sin(rotation) + oldX*Math.Cos(rotation);
                dynamic newY = oldY*Math.Cos(rotation) - oldX*Math.Sin(rotation);

                if (isSecondOrFourthQuadrant)
                {
                    newX = oldY * Math.Cos(rotation) - oldX * Math.Sin(rotation);
                    newY = oldY * Math.Sin(rotation) + oldX * Math.Cos(rotation);
                }

                _shape.Nodes.SetPosition(i, newX, newY);
            }

            _shape.Rotation = _originalRotation;

            UpdateAbsoluteWidth();
            UpdateAbsoluteHeight();

            VisualLeft = centerLeft - _absoluteWidth/2;
            VisualTop = centerTop - _absoluteHeight/2;
        }

        #endregion
    }
}
