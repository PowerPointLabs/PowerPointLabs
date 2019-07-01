using System;
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
            try
            {
                PPLClipboard.Instance.LockClipboard();
                slide.Copy();
                return shapes.PasteSpecial(PpPasteDataType.ppPastePNG)[1];
            }
            catch (Exception e)
            {
                throw e;
            }
            finally
            {
                PPLClipboard.Instance.ReleaseClipboard();
            }
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
