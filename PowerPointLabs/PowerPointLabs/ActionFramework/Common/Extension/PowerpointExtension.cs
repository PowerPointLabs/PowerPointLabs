using System;
using System.Collections.Generic;
using System.Reflection;
using Microsoft.Office.Interop.PowerPoint;
using PowerPointLabs.Models;
using PowerPointLabs.Utils;

namespace PowerPointLabs.ActionFramework.Common.Extension
{
    /// <summary>
    /// Contains extensions for <seealso cref="Shape"/> clipboard operations.
    /// </summary>
    public static class PowerpointExtension
    {
        // groups the shaperange only if there is 2 or more elements
        public static Shape SafeGroup(this ShapeRange range, PowerPointSlide slide)
        {
            if (range.Count > 1)
            {
                return SafeGroupPlaceholders(range, slide);
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

        public static Shape ConvertToNonPlaceHolder(this Shape shape, Shapes shapeSource)
        {
            if (shape.Type != Microsoft.Office.Core.MsoShapeType.msoPlaceholder)
            {
                return shape;
            }
            PowerPointLabs.SyncLab.ObjectFormats.Format[] formats = ShapeUtil.GetCopyableFormats(shape);
            Shape newShape = ShapeUtil.CopyMsoPlaceHolder(formats, shape, shapeSource);
            if (newShape == null)
            {
                throw new MsoPlaceholderException(shape.PlaceholderFormat.Type);
            }
            shape.SafeDelete();
            return newShape;
        }

        public static void SafeDelete(this ShapeRange shapeRange)
        {
            shapeRange.Delete();
            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(shapeRange);
            GC.Collect();
        }

        /// <summary>
        /// Releases all references to <seealso cref="Shape"/> before calling GC to collect.
        /// Required for protection against shape corruption from undo.
        /// </summary>
        /// <param name="shape"></param>
        public static void SafeDelete(this Shape shape)
        {
            shape.Delete();
            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(shape);
            GC.Collect();
        }

        public static void SafeSelect(this Shape shape)
        {
            // TODO
            //Globals.ThisAddIn.Application.ActiveWindow.View.GotoSlide(currentSlide.Index);
            throw new System.NotImplementedException();
        }

        // Warning: This method pastes each individually
        public static void SafeCopyPlaceholders(this Shapes shapes, ShapeRange shapeRange)
        {
            foreach (Shape s in shapeRange)
            {
                shapes.SafeCopyPlaceholder(s);
            }
        }

        // copies placeholder textboxes safely
        public static Shape SafeCopyPlaceholder(this Shapes shapes, Shape shape)
        {
            if (shape.Type != Microsoft.Office.Core.MsoShapeType.msoPlaceholder)
            {
                return shapes.SafeCopy(shape);
            }
            PowerPointLabs.SyncLab.ObjectFormats.Format[] formats = ShapeUtil.GetCopyableFormats(shape);
            Shape newShape = ShapeUtil.CopyMsoPlaceHolder(formats, shape, shapes);
            if (newShape == null)
            {
                throw new MsoPlaceholderException(shape.PlaceholderFormat.Type);
            }
            return newShape;
        }

        public static Shape SafeCopy(this Shapes shapes, Shape shape)
        {
            return PPLClipboard.Instance.LockAndRelease(() =>
            {
                shape.Copy();
                return shapes.Paste()[1];
            });
        }

        public static ShapeRange SafeCopy(this Shapes shapes, ShapeRange shapeRange)
        {
            return PPLClipboard.Instance.LockAndRelease(() =>
            {
                shapeRange.Copy();
                return shapes.Paste();
            });
        }

        public static Shape SafeCopySlide(this Shapes shapes, PowerPointSlide slide)
        {
            return PPLClipboard.Instance.LockAndRelease(() =>
            {
                slide.Copy();
                return shapes.PasteSpecial(PpPasteDataType.ppPastePNG)[1];
            });
        }

        public static Shape SafeCut(this Shapes shapes, Shape shape)
        {
            return PPLClipboard.Instance.LockAndRelease(() =>
            {
                shape.Cut();
                return shapes.PasteSpecial(PpPasteDataType.ppPastePNG)[1];
            });
        }

        public static Shape SafeCut(this Shapes shapes, ShapeRange selection)
        {
            return PPLClipboard.Instance.LockAndRelease(() =>
            {
                selection.Cut();
                return shapes.PasteSpecial(PpPasteDataType.ppPastePNG)[1];
            });
        }

        public static Shape SafeCopyPNG(this Shapes shapes, Shape shape)
        {
            return PPLClipboard.Instance.LockAndRelease(() =>
            {
                shape.Copy();
                return shapes.PasteSpecial(PpPasteDataType.ppPastePNG)[1];
            });
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

        // and also converts placeholders to lookalikes
        private static Shape SafeGroupPlaceholders(this ShapeRange range, PowerPointSlide slide)
        {
            List<Shape> finalShapes = new List<Shape>();
            foreach (Shape shape in range)
            {
                Shape finalShape = shape.ConvertToNonPlaceHolder(slide.Shapes);
                finalShapes.Add(finalShape);
            }
            return slide.ToShapeRange(finalShapes).Group();
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

        private static void TryAndCatch(Action action)
        {
            try
            {
                action();
            }
            catch (Exception e)
            {
                Log.Logger.LogException(e, action.Method.Name);
            }
        }
    }
}
