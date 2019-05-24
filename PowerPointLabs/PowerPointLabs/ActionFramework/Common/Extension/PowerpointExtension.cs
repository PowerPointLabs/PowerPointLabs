using System;
using System.Reflection;
using Microsoft.Office.Interop.PowerPoint;

namespace PowerPointLabs.ActionFramework.Common.Extension
{
    public static class PowerpointExtension
    {
        // groups the shaperange only if there is 2 or more elements
        public static Shape SafeGroup(this ShapeRange range)
        {
            if (range.Count > 1)
            {
                return range.Group();
            }
            else if (range.Count == 1)
            {
                return range[1];
            }
            else
            {
                return null;
            }
        }

        // copies placeholder textboxes safely
        public static Shape SafeCopy(this Shapes shapes, Shape shape)
        {
            if (shape.IsEmptyPlaceholder())
            {
                Shape newShape = shapes.AddTextbox(shape.TextFrame.Orientation, shape.Left, shape.Top, shape.Width, shape.Height);
                newShape.Fill.CopyPropertiesAndFieldsFrom(shape.Fill);
                newShape.Line.CopyPropertiesAndFieldsFrom(shape.Line);
                newShape.TextFrame.TextRange.Font.CopyPropertiesAndFieldsFrom(shape.TextFrame.TextRange.Font);
                return newShape;
            }
            shape.Copy();
            return shapes.Paste()[1];
        }

        // Referred to code in ColorsLabPaneWPFxaml, done manually
        public static void CopyColorFrom(this Shape targetShape, Shape sourceShape)
        {
            targetShape.Fill.ForeColor = sourceShape.Fill.ForeColor;
            targetShape.Fill.BackColor = sourceShape.Fill.BackColor;

            targetShape.Line.ForeColor = sourceShape.Line.ForeColor;
            targetShape.Line.BackColor = sourceShape.Line.BackColor;
            targetShape.Line.Visible = sourceShape.Line.Visible;

            if (sourceShape.TextFrame.HasText == Microsoft.Office.Core.MsoTriState.msoTrue)
            {
                Font newShapeFont = targetShape.TextFrame.TextRange.Font;
                Font shapeFont = sourceShape.TextFrame.TextRange.Font;
                newShapeFont.CopyFontFrom(shapeFont);
            }
        }

        public static void CopyPropertiesAndFieldsFrom<T>(this T target, T source)
        {
            target.CopyPropertiesFrom(source);
            target.CopyFieldsFrom(source);
        }

        // copies properties recursively. Does not detect loops
        public static void CopyPropertiesFrom<T>(this T target, T source)
        {
            typeof(T).CopyProperties(target, source);
        }

        public static void CopyFieldsFrom<T>(this T target, T source)
        {
            typeof(T).CopyFields(target, source);
        }

        public static bool IsEmptyPlaceholder(this Shape shape)
        {
            return shape.Type == Microsoft.Office.Core.MsoShapeType.msoPlaceholder &&
                shape.TextFrame.TextRange.Text.Length == 0;
        }

        private static void CopyPropertiesAndFields(this Type t, object target, object source)
        {
            t.CopyProperties(target, source);
            t.CopyFields(target, source);
        }

        private static void CopyProperties(this Type type, object target, object source)
        {
            foreach (PropertyInfo i in type.GetProperties())
            {
                Type t = i.PropertyType;
                if (!i.CanRead)
                {
                    continue; // can't copy at all
                }
                else if (!i.CanWrite)
                {
                    // recursive copy
                    TryAndCatch(() => t.CopyPropertiesAndFields(i.GetValue(target), i.GetValue(source)));
                    continue;
                }
                // direct copy
                TryAndCatch(() => i.SetValue(target, i.GetValue(source)));
            }
        }

        private static void CopyFields(this Type t, object target, object source)
        {
            foreach (FieldInfo i in t.GetFields())
            {
                TryAndCatch(() => i.SetValue(target, i.GetValue(source)));
            }
        }

        private static void TryAndCatch(Action a)
        {
            try
            {
                a();
            }
            catch
            {
                // do nothing
            }
        }

        // a brute force copy
        private static void CopyFontFrom(this Font font, Font other)
        {
            font.AutoRotateNumbers = other.AutoRotateNumbers;
            font.BaselineOffset = other.BaselineOffset;
            font.Bold = other.Bold;
            font.Emboss = other.Emboss;
            font.Italic = other.Italic;
            font.Shadow = other.Shadow;
            font.Subscript = other.Subscript;
            font.Superscript = other.Superscript;
            font.Underline = other.Underline;
            font.Color.CopyColorFrom(other.Color);
        }

        private static void CopyColorFrom(this ColorFormat color, ColorFormat other)
        {
            color.Brightness = other.Brightness;
            color.RGB = other.RGB;
            color.SchemeColor = other.SchemeColor;
            color.TintAndShade = other.TintAndShade;
        }
    }
}
