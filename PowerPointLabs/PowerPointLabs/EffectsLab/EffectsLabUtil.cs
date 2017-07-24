using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Windows.Forms;

using Microsoft.Office.Interop.PowerPoint;
using PowerPointLabs.Models;
using PowerPointLabs.Utils;
using PowerPointLabs.Views;

using Office = Microsoft.Office.Core;

namespace PowerPointLabs.EffectsLab
{
    public static class EffectsLabUtil
    {
        private const string MessageBoxTitle = "Error";
        private const string ErrorMessageNoSelection = TextCollection.EffectsLabBlurSelectedErrorNoSelection;
        private const string ErrorMessageNonShapeOrTextBox = TextCollection.EffectsLabBlurSelectedErrorNonShapeOrTextBox;

        internal static ShapeRange UngroupAllShapeRange(PowerPointSlide curSlide, ShapeRange shapeRange)
        {
            List<Shape> originalShapeList = new List<Shape>();
            List<Shape> ungroupedShapeList = new List<Shape>();

            for (int i = 1; i <= shapeRange.Count; i++)
            {
                var shape = shapeRange[i];
                if (ShapeUtil.IsCorrupted(shape))
                {
                    shape = ShapeUtil.CorruptionCorrection(shape, curSlide);
                }
                originalShapeList.Add(shape);
            }

            for (int i = 0; i < originalShapeList.Count; i++)
            {
                if (originalShapeList[i].Type == Office.MsoShapeType.msoGroup)
                {
                    var subRange = originalShapeList[i].Ungroup();
                    foreach (Shape item in subRange)
                    {
                        originalShapeList.Add(item);
                    }
                }
                else if (originalShapeList[i].Type == Office.MsoShapeType.msoPlaceholder ||
                    originalShapeList[i].Type == Office.MsoShapeType.msoTextBox ||
                    originalShapeList[i].Type == Office.MsoShapeType.msoAutoShape ||
                    originalShapeList[i].Type == Office.MsoShapeType.msoFreeform)
                {
                    ungroupedShapeList.Add(originalShapeList[i]);
                }
                else
                {
                    throw new Exception(ErrorMessageNonShapeOrTextBox);
                }
            }

            var ungroupedShapeRange = curSlide.ToShapeRange(ungroupedShapeList);

            return ungroupedShapeRange;
        }

        internal static PowerPointBgEffectSlide GenerateEffectSlide(PowerPointSlide curSlide, Selection selection, bool generateOnRemainder)
        {
            PowerPointSlide dupSlide = null;

            try
            {
                ShapeRange shapeRange = selection.ShapeRange;
                if (selection.HasChildShapeRange)
                {
                    shapeRange = selection.ChildShapeRange;
                }

                if (shapeRange.Count != 0)
                {
                    dupSlide = curSlide.Duplicate();
                }

                shapeRange.Cut();

                var effectSlide = PowerPointBgEffectSlide.BgEffectFactory(curSlide.GetNativeSlide(), generateOnRemainder);

                if (dupSlide != null)
                {
                    if (generateOnRemainder)
                    {
                        dupSlide.Delete();
                    }
                    else
                    {
                        dupSlide.MoveTo(curSlide.Index);
                        curSlide.Delete();
                    }
                }

                return effectSlide;
            }
            catch (InvalidOperationException e)
            {
                MessageBox.Show(e.Message);
                return null;
            }
            catch (COMException)
            {
                if (dupSlide != null)
                {
                    dupSlide.Delete();
                }

                MessageBox.Show("Please select at least 1 shape");
                return null;
            }
            catch (Exception e)
            {
                if (dupSlide != null)
                {
                    dupSlide.Delete();
                }

                ErrorDialogBox.ShowDialog("Error", e.Message, e);
                return null;
            }
        }

        internal static Shape DuplicateShapeInPlace(Shape shape)
        {
            var duplicateShape = shape.Duplicate()[1];
            duplicateShape.Left = shape.Left;
            duplicateShape.Top = shape.Top;

            var match = System.Text.RegularExpressions.Regex.Match(duplicateShape.Name, @"\d+$");
            if (!match.Success || int.Parse(match.Value) != duplicateShape.Id - 1)
            {
                duplicateShape.Name += " " + (duplicateShape.Id - 1);
            }

            return duplicateShape;
        }

        internal static void ShowErrorMessageBox(string content, Exception exception = null)
        {
            if (exception == null
                || content.Equals(ErrorMessageNoSelection)
                || content.Equals(ErrorMessageNonShapeOrTextBox))
            {
                MessageBox.Show(content, MessageBoxTitle);
            }
            else
            {
                ErrorDialogBox.ShowDialog(MessageBoxTitle, content, exception);
            }
        }

        internal static void ShowNoSelectionErrorMessage()
        {
            ShowErrorMessageBox(ErrorMessageNoSelection);
        }

        internal static void ShowIncorrectSelectionErrorMessage()
        {
            ShowErrorMessageBox(ErrorMessageNonShapeOrTextBox);
        }
    }
}
