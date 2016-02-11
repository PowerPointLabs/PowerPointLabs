using System;
using System.Drawing;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointLabs.ResizeLab
{
    internal class ResizeLabShape
    {
        private readonly PowerPoint.Shape _shape;
        private float _virtualWidth;
        private float _virtualHeight;

        public ResizeLabShape(PowerPoint.Shape shape)
        {
            _shape = shape;

            UpdateVirtualWidth();
            UpdateVirtualHeight();
        }

        public float ActualWidth
        {
            get { return _shape.Width; }
            set
            {
                _shape.Width = value;
                UpdateVirtualWidth();
            }
        }

        public float ActualHeight
        {
            get { return _shape.Height; }
            set
            {
                _shape.Height = value;
                UpdateVirtualHeight();
            }
        }

        public float VirtualWidth
        {
            get { return _virtualWidth; }
            set
            {
                _virtualWidth = value;
                SetToVirtualDimension();
            }
        }

        public float VirtualHeight
        {
            get { return _virtualHeight; }
            set
            {
                _virtualHeight = value;
                SetToVirtualDimension();
            }
        }

        public float Top
        {
            get { return _shape.Top; }
            set { _shape.Top = value; }
        }

        public float Left
        {
            get { return _shape.Left; }
            set { _shape.Left = value; }
        }

        /// <summary>
        /// Update the virtual width according to the actual shape width and height.
        /// </summary>
        private void UpdateVirtualWidth()
        {
            var rotation = GetStandardizedRotation(_shape.Rotation);

            if (IsInQuadrant(_shape.Rotation))
            {
                _virtualWidth = (float) (_shape.Height*Math.Sin(rotation) + _shape.Width*Math.Cos(rotation));
            }
            else if ((int) _shape.Rotation == 90 || (int) _shape.Rotation == 270)
            {
                _virtualWidth = _shape.Height;
            }
            else
            {
                _virtualWidth = _shape.Width;
            }
        }

        /// <summary>
        /// Update the virtual height according to the actual shape width and height.
        /// </summary>
        private void UpdateVirtualHeight()
        {
            var rotation = GetStandardizedRotation(_shape.Rotation);

            if (IsInQuadrant(_shape.Rotation))
            {
                _virtualHeight = (float) (_shape.Height*Math.Cos(rotation) + _shape.Width*Math.Sin(rotation));
            }
            else if ((int) _shape.Rotation == 90 || (int) _shape.Rotation == 270)
            {
                _virtualHeight = _shape.Width;
            }
            else
            {
                _virtualHeight = _shape.Height;
            }
        }

        /// <summary>
        /// Set the actual width and height according to the virtual dimension (e.g. width and height).
        /// </summary>
        private void SetToVirtualDimension()
        {
            var rotation = GetStandardizedRotation(_shape.Rotation);
            var sinAngle = Math.Sin(rotation);
            var cosAngle = Math.Cos(rotation);
            var ratio = sinAngle/cosAngle;

            _shape.Height = (float) ((_virtualWidth*ratio - _virtualHeight)/(sinAngle*ratio - cosAngle));
            _shape.Width = (float) ((_virtualWidth - _shape.Height*sinAngle)/cosAngle);
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
