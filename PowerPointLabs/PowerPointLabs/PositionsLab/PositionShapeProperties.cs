using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Core;

namespace PowerPointLabs.PositionsLab
{
    public class PositionShapeProperties
    {
        private float _rotation;
        private System.Drawing.PointF _position;
        private MsoTriState _flipHorizontalState;
        private MsoTriState _flipVerticalState;

        public PositionShapeProperties(System.Drawing.PointF position, float rotation, MsoTriState flipHorizontalState, MsoTriState flipVerticalState)
        {
            this._position = position;
            this._rotation = rotation;
            this._flipHorizontalState = flipHorizontalState;
            this._flipVerticalState = flipVerticalState;
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

        public MsoTriState FlipHorizontalState
        {
            get { return _flipHorizontalState; }
            set { _flipHorizontalState = value; }
        }

        public MsoTriState FlipVerticalState
        {
            get { return _flipVerticalState; }
            set { _flipVerticalState = value; }
        }
    }
}
