using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace PowerPointLabs.PositionsLab
{
    public class PositionShapeProperties
    {
        private float _rotation;
        private System.Drawing.PointF _position;

        public PositionShapeProperties(System.Drawing.PointF position, float rotation)
        {
            this._position = position;
            this._rotation = rotation;
        }

        public System.Drawing.PointF Position
        {
            get { return _position; }
            set { _position = value; }
        }

        public float Rotation
        {
            get { return _rotation; }
            set { _rotation = value; }
        }
    }
}
