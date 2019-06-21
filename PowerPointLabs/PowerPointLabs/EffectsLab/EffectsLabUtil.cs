using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Windows.Forms;

using Microsoft.Office.Interop.PowerPoint;

using PowerPointLabs.Models;
using PowerPointLabs.TextCollection;
using PowerPointLabs.Utils;
using PowerPointLabs.Utils.Windows;
using PowerPointLabs.Views;

using Office = Microsoft.Office.Core;

namespace PowerPointLabs.EffectsLab
{
    public static class EffectsLabUtil
    {
        internal static ShapeRange UngroupAllShapeRange(PowerPointSlide curSlide, ShapeRange shapeRange)
        {
            List<Shape> originalShapeList = new List<Shape>();
            List<Shape> ungroupedShapeList = new List<Shape>();

            for (int i = 1; i <= shapeRange.Count; i++)
            {
                Shape shape = shapeRange[i];
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
                    ShapeRange subRange = originalShapeList[i].Ungroup();
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
                    throw new Exception(EffectsLabText.ErrorBlurSelectedNonShapeOrTextBox);
                }
            }

            ShapeRange ungroupedShapeRange = curSlide.ToShapeRange(ungroupedShapeList);

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

                PowerPointBgEffectSlide effectSlide = PowerPointBgEffectSlide.BgEffectFactory(curSlide.GetNativeSlide(), generateOnRemainder);

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

                MessageBox.Show(TextCollection.EffectsLabText.ErrorSelectAtLeastOneShape);
                return null;
            }
            catch (Exception e)
            {
                if (dupSlide != null)
                {
                    dupSlide.Delete();
                }

                ErrorDialogBox.ShowDialog(CommonText.ErrorTitle, e.Message, e);
                return null;
            }
        }

        internal static Shape DuplicateShapeInPlace(Shape shape)
        {
            Shape duplicateShape = shape.Duplicate()[1];
            duplicateShape.Left = shape.Left;
            duplicateShape.Top = shape.Top;

            System.Text.RegularExpressions.Match match = System.Text.RegularExpressions.Regex.Match(duplicateShape.Name, @"\d+$");
            if (!match.Success || int.Parse(match.Value) != duplicateShape.Id - 1)
            {
                duplicateShape.Name += " " + (duplicateShape.Id - 1);
            }

            return duplicateShape;
        }

        internal static void ShowErrorMessageBox(string content, Exception exception = null)
        {
            if (exception == null
                || content.Equals(EffectsLabText.ErrorBlurSelectedNoSelection)
                || content.Equals(EffectsLabText.ErrorBlurSelectedNonShapeOrTextBox))
            {
                MessageBoxUtil.Show(content, EffectsLabText.ErrorDialogTitle);
            }
            else
            {
                ErrorDialogBox.ShowDialog(EffectsLabText.ErrorDialogTitle, content, exception);
            }
        }

        internal static void ShowNoSelectionErrorMessage()
        {
            ShowErrorMessageBox(EffectsLabText.ErrorBlurSelectedNoSelection);
        }

        internal static void ShowIncorrectSelectionErrorMessage()
        {
            ShowErrorMessageBox(EffectsLabText.ErrorBlurSelectedNonShapeOrTextBox);
        }
    }
}
