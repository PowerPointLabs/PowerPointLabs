using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace PowerPointLabs.ResizeLab
{
    internal class ShapeProperties
    {
        public string Name { get; }
        public float Top { get; }
        public float Left { get; }
        public float Width { get; }
        public float Height { get; }

        public ShapeProperties(string name, float top, float left, float width, float height)
        {
            Name = name;
            Top = top;
            Left = left;
            Width = width;
            Height = height;
        }
    }
}
