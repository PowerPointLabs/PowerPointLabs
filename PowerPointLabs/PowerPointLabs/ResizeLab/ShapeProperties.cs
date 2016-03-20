using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace PowerPointLabs.ResizeLab
{
    public class ShapeProperties
    {
        public string Name { get; }
        public int Id { get; }
        public float Top { get; }
        public float Left { get; }
        public float AbsoluteWidth { get; }
        public float AbsoluteHeight { get; }
        public float ShapeRotation { get; }

        public ShapeProperties(int id, float top, float left, float absoluteWidth, float absoluteHeight, float rotation)
        {
            Id = id;
            Top = top;
            Left = left;
            AbsoluteWidth = absoluteWidth;
            AbsoluteHeight = absoluteHeight;
            ShapeRotation = rotation;
        }

        public ShapeProperties(string name, float top, float left, float absoluteWidth, float absoluteHeight)
        {
            Name = name;
            Top = top;
            Left = left;
            AbsoluteWidth = absoluteWidth;
            AbsoluteHeight = absoluteHeight;
        }
    }
}
