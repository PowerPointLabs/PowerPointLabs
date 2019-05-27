using System;
using System.Collections.Generic;
using System.Linq;
using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.ActionFramework.Common.Log;
using PowerPointLabs.Models;
using PowerPointLabs.Utils;
using PowerPointLabs.Views;

using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointLabs.ZoomLab
{
    internal static class AutoZoom
    {
#pragma warning disable 0618
        public static void AddDrillDownAnimation()
        {
            if (!IsSelectingShapes())
            {
                return;
            }

            AddDrillDownAnimation(Globals.ThisAddIn.Application.ActiveWindow.Selection.ShapeRange[1],
                PowerPointCurrentPresentationInfo.CurrentSlide);
        }

        public static void AddDrillDownAnimation(PowerPoint.Shape selectedShape, PowerPointSlide currentSlide)
        {
            PowerPointDrillDownSlide addedSlide;
            AddDrillDownAnimation(selectedShape, currentSlide, out addedSlide);
        }

        public static void AddDrillDownAnimation(PowerPoint.Shape selectedShape, PowerPointSlide currentSlide,
            out PowerPointDrillDownSlide addedSlide, bool includeAckSlide = true, bool deletePreviouslyAdded = true)
        {
            try
            {
                if (currentSlide == null || currentSlide.Index == PowerPointPresentation.Current.SlideCount)
                {
                    System.Windows.Forms.MessageBox.Show(TextCollection.ZoomLabText.ErrorInvalidNextSlide, TextCollection.ZoomLabText.ErrorUnableToAddAnimationsCaption);
                    addedSlide = null;
                    return;
                }

                //Pick up the border and shadow style, to be applied to zoomed shape
                selectedShape.PickUp();
                PrepareZoomShape(currentSlide, ref selectedShape);
                PowerPointSlide nextSlide = GetNextSlide(currentSlide, deletePreviouslyAdded);

                PowerPoint.Shape nextSlidePicture = null, shapeToZoom = null;

                currentSlide.HideIndicator();
                if (ZoomLabSettings.BackgroundZoomChecked)
                {
                    nextSlidePicture = GetNextSlidePictureWithBackground(currentSlide, nextSlide);
                    nextSlidePicture.Apply();
                    PrepareNextSlidePicture(currentSlide, selectedShape, ref nextSlidePicture);

                    addedSlide = (PowerPointDrillDownSlide)currentSlide.CreateDrillDownSlide();
                    addedSlide.DeleteAllShapes();

                    shapeToZoom = addedSlide.Shapes.SafeCopyPlaceholder(nextSlidePicture);
                    addedSlide.DeleteShapeAnimations(shapeToZoom);

                    PowerPoint.Shape backgroundShape = AddSlideAsShape(currentSlide, addedSlide);
                    backgroundShape.Apply();
                    ShapeUtil.FitShapeToSlide(ref backgroundShape);
                    backgroundShape.ZOrder(Office.MsoZOrderCmd.msoSendBackward);
                    backgroundShape.Name = "PPTZoomInShape" + DateTime.Now.ToString("yyyyMMddHHmmssffff");

                    addedSlide.PrepareForDrillDown();
                    addedSlide.AddDrillDownAnimationBackground(backgroundShape, shapeToZoom, nextSlidePicture);
                }
                else
                {
                    PowerPoint.Shape pictureOnNextSlide = null;
                    nextSlidePicture = GetNextSlidePictureWithoutBackground(currentSlide, nextSlide, out pictureOnNextSlide);
                    nextSlidePicture.Apply();
                    PrepareNextSlidePicture(currentSlide, selectedShape, ref nextSlidePicture);

                    addedSlide = (PowerPointDrillDownSlide)currentSlide.CreateDrillDownSlide();
                    addedSlide.DeleteAllShapes();

                    shapeToZoom = addedSlide.Shapes.SafeCopyPlaceholder(nextSlidePicture);
                    addedSlide.DeleteShapeAnimations(shapeToZoom);

                    PowerPoint.Shape backgroundShape = AddSlideAsShape(currentSlide, addedSlide);
                    backgroundShape.Apply();
                    ShapeUtil.FitShapeToSlide(ref backgroundShape);
                    backgroundShape.ZOrder(Office.MsoZOrderCmd.msoSendBackward);
                    backgroundShape.Name = "PPTZoomInShape" + DateTime.Now.ToString("yyyyMMddHHmmssffff");

                    addedSlide.PrepareForDrillDown();
                    addedSlide.AddDrillDownAnimationNoBackground(backgroundShape, shapeToZoom, pictureOnNextSlide);
                    pictureOnNextSlide.Delete();
                }
                currentSlide.ShowIndicator();

                Globals.ThisAddIn.Application.ActiveWindow.View.GotoSlide(addedSlide.Index);
                Globals.ThisAddIn.Application.CommandBars.ExecuteMso("AnimationPreview");
                if (includeAckSlide)
                {
                    PowerPointPresentation.Current.AddAckSlide();
                }
            }
            catch (Exception e)
            {
                Logger.LogException(e, "AddDrillDownAnimation");
                ErrorDialogBox.ShowDialog("Error when adding drill down animation", "An error occurred when adding drill down animation.", e);
                throw;
            }
        }

        public static void AddStepBackAnimation()
        {
            if (!IsSelectingShapes())
            {
                return;
            }

            AddStepBackAnimation(Globals.ThisAddIn.Application.ActiveWindow.Selection.ShapeRange[1],
                PowerPointCurrentPresentationInfo.CurrentSlide);
        }

        public static void AddStepBackAnimation(PowerPoint.Shape selectedShape, PowerPointSlide currentSlide)
        {
            PowerPointStepBackSlide addedSlide;
            AddStepBackAnimation(selectedShape, currentSlide, out addedSlide);
        }

        public static void AddStepBackAnimation(PowerPoint.Shape selectedShape, PowerPointSlide currentSlide,
            out PowerPointStepBackSlide addedSlide, bool includeAckSlide = true, bool deletePreviouslyAdded = true)
        {
            try
            {
                if (currentSlide == null || currentSlide.Index == 1)
                {
                    System.Windows.Forms.MessageBox.Show(TextCollection.ZoomLabText.ErrorInvalidPreviousSlide, TextCollection.ZoomLabText.ErrorUnableToAddAnimationsCaption);
                    addedSlide = null;
                    return;
                }

                //Pick up the border and shadow style, to be applied to zoomed shape
                selectedShape.PickUp();
                PrepareZoomShape(currentSlide, ref selectedShape);
                PowerPointSlide previousSlide = GetPreviousSlide(currentSlide, deletePreviouslyAdded);

                PowerPoint.Shape previousSlidePicture = null, shapeToZoom = null;
                
                currentSlide.HideIndicator();
                if (ZoomLabSettings.BackgroundZoomChecked)
                {
                    previousSlidePicture = GetPreviousSlidePictureWithBackground(currentSlide, previousSlide);
                    previousSlidePicture.Apply();
                    PreparePreviousSlidePicture(selectedShape, ref previousSlidePicture);

                    addedSlide = (PowerPointStepBackSlide)previousSlide.CreateStepBackSlide();
                    addedSlide.DeleteAllShapes();

                    PowerPoint.Shape backgroundShape = null;
                    shapeToZoom = GetStepBackWithBackgroundShapeToZoom(currentSlide, addedSlide, previousSlidePicture, out backgroundShape);
                    shapeToZoom.Apply();

                    addedSlide.PrepareForStepBack();
                    addedSlide.AddStepBackAnimationBackground(shapeToZoom, backgroundShape, previousSlidePicture);
                }
                else
                {
                    addedSlide = (PowerPointStepBackSlide)previousSlide.CreateStepBackSlide();
                    addedSlide.DeleteAllShapes();

                    shapeToZoom = GetStepBackWithoutBackgroundShapeToZoom(addedSlide, previousSlide);
                    shapeToZoom.Apply();
                    previousSlidePicture = currentSlide.Shapes.SafeCopyPNG(shapeToZoom);
                    previousSlidePicture.Apply();
                    PreparePreviousSlidePicture(selectedShape, ref previousSlidePicture);

                    addedSlide.PrepareForStepBack();
                    addedSlide.AddStepBackAnimationNonBackground(shapeToZoom, previousSlidePicture);
                }
                currentSlide.ShowIndicator();

                currentSlide.Transition.EntryEffect = PowerPoint.PpEntryEffect.ppEffectFadeSmoothly;
                currentSlide.Transition.Duration = 0.25f;
                Globals.ThisAddIn.Application.ActiveWindow.View.GotoSlide(addedSlide.Index);
                Globals.ThisAddIn.Application.CommandBars.ExecuteMso("AnimationPreview");
                if (includeAckSlide)
                {
                    PowerPointPresentation.Current.AddAckSlide();
                }
            }
            catch (Exception e)
            {
                Logger.LogException(e, "AddStepBackAnimation");
                ErrorDialogBox.ShowDialog("Error when adding step back animation", "An error occurred when adding step back animation.", e);
                throw;
            }
        }

        private static void RemoveTextFromShape(PowerPoint.Shape shape)
        {
            if (shape.HasTextFrame == Office.MsoTriState.msoTrue && shape.TextFrame.HasText == Office.MsoTriState.msoTrue)
            {
                shape.TextFrame.TextRange.Text = "";
            }
        }

        private static void PrepareZoomShape(PowerPointSlide currentSlide, ref PowerPoint.Shape selectedShape)
        {
            currentSlide.DeleteShapeAnimations(selectedShape);
            RemoveTextFromShape(selectedShape);
            selectedShape.Rotation = 0;
        }

        //Delete previously added drill down slides
        private static PowerPointSlide GetNextSlide(PowerPointSlide currentSlide, bool deletePreviouslyAdded = true)
        {
            PowerPointSlide nextSlide = PowerPointPresentation.Current.Slides[currentSlide.Index];
            PowerPointSlide tempSlide = nextSlide;
            while (nextSlide.Name.Contains("PPTLabsZoomIn") && nextSlide.Index < PowerPointPresentation.Current.SlideCount)
            {
                nextSlide = PowerPointPresentation.Current.Slides[tempSlide.Index];
                if (deletePreviouslyAdded)
                {
                    tempSlide.Delete();
                }

                tempSlide = nextSlide;
            }
            nextSlide.Transition.EntryEffect = PowerPoint.PpEntryEffect.ppEffectFadeSmoothly;
            nextSlide.Transition.Duration = 0.25f;
            return nextSlide;
        }

        //Delete previously added step back slides
        private static PowerPointSlide GetPreviousSlide(PowerPointSlide currentSlide, bool deletePreviouslyAdded = true)
        {
            PowerPointSlide previousSlide = PowerPointPresentation.Current.Slides[currentSlide.Index - 2];
            PowerPointSlide tempSlide = previousSlide;
            while (previousSlide.Name.Contains("PPTLabsZoomOut") && previousSlide.Index > 1)
            {
                previousSlide = PowerPointPresentation.Current.Slides[tempSlide.Index - 2];
                if (deletePreviouslyAdded)
                {
                    tempSlide.Delete();
                }

                tempSlide = previousSlide;
            }

            return previousSlide;
        }

        //Return picture copy of next slide where shapes with exit animations have been deleted
        private static PowerPoint.Shape GetNextSlidePictureWithoutBackground(PowerPointSlide currentSlide, PowerPointSlide nextSlide, out PowerPoint.Shape pictureOnNextSlide)
        {
            Globals.ThisAddIn.Application.ActiveWindow.Selection.Unselect();
            Globals.ThisAddIn.Application.ActiveWindow.View.GotoSlide(nextSlide.Index);

            List<PowerPoint.Shape> copiedShapes = new List<PowerPoint.Shape>();
            foreach (PowerPoint.Shape sh in nextSlide.Shapes)
            {
                if (!nextSlide.HasEntryAnimation(sh) && !ShapeUtil.IsHidden(sh))
                {
                    PowerPoint.Shape shapeCopy = currentSlide.Shapes.SafeCopyPlaceholder(sh);
                    LegacyShapeUtil.CopyShapeAttributes(sh, ref shapeCopy);
                    copiedShapes.Add(shapeCopy);
                }
            }

            Globals.ThisAddIn.Application.ActiveWindow.View.GotoSlide(currentSlide.Index);
            SelectAllShape(copiedShapes);
            PowerPoint.Selection sel = Globals.ThisAddIn.Application.ActiveWindow.Selection;
            PowerPoint.Shape shapeGroup = sel.ShapeRange.SafeGroup();

            pictureOnNextSlide = nextSlide.Shapes.SafeCopyPNG(shapeGroup);
            LegacyShapeUtil.CopyShapePosition(shapeGroup, ref pictureOnNextSlide);
            shapeGroup.Delete();

            PowerPoint.Shape slidePicture = currentSlide.Shapes.SafeCopyPNG(pictureOnNextSlide);
            return slidePicture;
        }

        //Return picture copy of next slide where shapes with exit animations have been deleted
        private static PowerPoint.Shape GetNextSlidePictureWithBackground(PowerPointSlide currentSlide, PowerPointSlide nextSlide)
        {
            PowerPointSlide nextSlideCopy = nextSlide.Duplicate();
            List<PowerPoint.Shape> shapes = nextSlideCopy.Shapes.Cast<PowerPoint.Shape>().ToList();
            IEnumerable<PowerPoint.Shape> matchingShapes = shapes.Where(current => nextSlideCopy.HasEntryAnimation(current));
            foreach (PowerPoint.Shape s in matchingShapes)
            {
                s.Delete();
            }
            PowerPoint.Shape slidePicture = AddSlideAsShape(nextSlideCopy, currentSlide);
            nextSlideCopy.Delete();
            return slidePicture;
        }

        //Return picture copy of previous slide where shapes with exit animations have been deleted
        private static PowerPoint.Shape GetPreviousSlidePictureWithBackground(PowerPointSlide currentSlide, PowerPointSlide previousSlide)
        {
            PowerPointSlide previousSlideCopy = previousSlide.Duplicate();
            List<PowerPoint.Shape> shapes = previousSlideCopy.Shapes.Cast<PowerPoint.Shape>().ToList();
            IEnumerable<PowerPoint.Shape> matchingShapes = shapes.Where(current => previousSlideCopy.HasExitAnimation(current));
            foreach (PowerPoint.Shape s in matchingShapes)
            {
                s.Delete();
            }
            PowerPoint.Shape slidePicture = AddSlideAsShape(previousSlideCopy, currentSlide);
            previousSlideCopy.Delete();
            return slidePicture;
        }

        //Set position, size and animations of the next slide copy
        private static void PrepareNextSlidePicture(PowerPointSlide currentSlide, PowerPoint.Shape selectedShape, ref PowerPoint.Shape nextSlidePicture)
        {
            nextSlidePicture.LockAspectRatio = Office.MsoTriState.msoTrue;
            if (selectedShape.Width > selectedShape.Height)
            {
                nextSlidePicture.Height = selectedShape.Height;
            }
            else
            {
                nextSlidePicture.Width = selectedShape.Width;
            }

            LegacyShapeUtil.CopyShapePosition(selectedShape, ref nextSlidePicture);
 
            selectedShape.Visible = Office.MsoTriState.msoFalse;
            nextSlidePicture.Name = "PPTZoomInShape" + DateTime.Now.ToString("yyyyMMddHHmmssffff");

            PowerPoint.Effect effectAppear = currentSlide.TimeLine.MainSequence.AddEffect(nextSlidePicture, PowerPoint.MsoAnimEffect.msoAnimEffectFade, PowerPoint.MsoAnimateByLevel.msoAnimateLevelNone, PowerPoint.MsoAnimTriggerType.msoAnimTriggerOnPageClick);
            effectAppear.Timing.Duration = 0.50f;
        }

        //Set position, size and animations of the previous slide copy
        private static void PreparePreviousSlidePicture(PowerPoint.Shape selectedShape, ref PowerPoint.Shape previousSlidePicture)
        {
            previousSlidePicture.LockAspectRatio = Office.MsoTriState.msoTrue;
            if (selectedShape.Width > selectedShape.Height)
            {
                previousSlidePicture.Height = selectedShape.Height;
            }
            else
            {
                previousSlidePicture.Width = selectedShape.Width;
            }

            LegacyShapeUtil.CopyShapePosition(selectedShape, ref previousSlidePicture);

            selectedShape.Visible = Office.MsoTriState.msoFalse;
            previousSlidePicture.Name = "PPTZoomOutShape" + DateTime.Now.ToString("yyyyMMddHHmmssffff");
        }

        private static bool IsSelectingShapes()
        {
            PowerPoint.Selection selection = Globals.ThisAddIn.Application.ActiveWindow.Selection;
            return selection.Type == PowerPoint.PpSelectionType.ppSelectionShapes && selection.ShapeRange.Count > 0;
        }


        private static PowerPoint.Shape GetStepBackWithBackgroundShapeToZoom(PowerPointSlide currentSlide, PowerPointSlide addedSlide, PowerPoint.Shape previousSlidePicture, out PowerPoint.Shape backgroundShape)
        {
            PowerPoint.Shape currentSlideCopy = AddSlideAsShape(currentSlide, addedSlide);

            ShapeUtil.FitShapeToSlide(ref currentSlideCopy);
            currentSlideCopy.Name = "PPTZoomOutShape" + DateTime.Now.ToString("yyyyMMddHHmmssffff");

            PowerPoint.Shape previousSlidePictureCopy = addedSlide.Shapes.SafeCopyPlaceholder(previousSlidePicture);

            Globals.ThisAddIn.Application.ActiveWindow.View.GotoSlide(addedSlide.Index);

            // Scale everything up by this ratio.
            float ratio = PowerPointPresentation.Current.SlideWidth / previousSlidePictureCopy.Width;

            currentSlideCopy.Width *= ratio;
            currentSlideCopy.Height *= ratio;
            currentSlideCopy.Left = -ratio * previousSlidePictureCopy.Left;
            currentSlideCopy.Top = -ratio * previousSlidePictureCopy.Top;

            // for some reason height is locked to width, so we only need to change width here.
            previousSlidePictureCopy.Width *= ratio;
            previousSlidePictureCopy.Left = 0;
            previousSlidePictureCopy.Top = 0;

            backgroundShape = currentSlideCopy;
            return previousSlidePictureCopy;
        }

        private static PowerPoint.Shape GetStepBackWithoutBackgroundShapeToZoom(PowerPointSlide addedSlide, PowerPointSlide previousSlide)
        {
            Globals.ThisAddIn.Application.ActiveWindow.Selection.Unselect();
            Globals.ThisAddIn.Application.ActiveWindow.View.GotoSlide(addedSlide.Index);

            List<PowerPoint.Shape> copiedShapes = new List<PowerPoint.Shape>();
            foreach (PowerPoint.Shape sh in previousSlide.Shapes)
            {
                if (!previousSlide.HasExitAnimation(sh) && !ShapeUtil.IsHidden(sh))
                {
                    PowerPoint.Shape shapeCopy = addedSlide.Shapes.SafeCopyPlaceholder(sh);
                    LegacyShapeUtil.CopyShapeAttributes(sh, ref shapeCopy);
                    copiedShapes.Add(shapeCopy);
                } 
            }

            Globals.ThisAddIn.Application.ActiveWindow.View.GotoSlide(addedSlide.Index);
            SelectAllShape(copiedShapes);
            PowerPoint.Selection sel = Globals.ThisAddIn.Application.ActiveWindow.Selection;
            PowerPoint.Shape shapeGroup = sel.ShapeRange.SafeGroup();

            PowerPoint.Shape previousSlidePicture = addedSlide.Shapes.SafeCopyPNG(shapeGroup);
            LegacyShapeUtil.CopyShapePosition(shapeGroup, ref previousSlidePicture);
            previousSlidePicture.Name = "PPTZoomOutShape" + DateTime.Now.ToString("yyyyMMddHHmmssffff");
            shapeGroup.Delete();

            return previousSlidePicture;
        }

        private static void SelectAllShape(List<PowerPoint.Shape> shapes)
        {
            foreach (PowerPoint.Shape shapeCopy in shapes)
            {
                shapeCopy.Select(Office.MsoTriState.msoFalse);
            }
        }

        private static PowerPoint.Shape AddSlideAsShape(PowerPointSlide slideToAdd, PowerPointSlide targetSlide)
        {
            try
            {
                // Export the slide as .png to a temporary location, then add it to shapes.
                // This yields a higher quality image compared to copy-pasting slide as image.
                string tempFilePath = FileDir.GetTemporaryPngFilePath();
                GraphicsUtil.ExportSlide(slideToAdd, tempFilePath);
                PowerPoint.Shape slideAsShape = targetSlide.Shapes.AddPicture2(tempFilePath,
                                                                                 Microsoft.Office.Core.MsoTriState.msoFalse,
                                                                                 Microsoft.Office.Core.MsoTriState.msoTrue,
                                                                                 0,
                                                                                 0);
                try
                {
                    FileDir.DeleteFile(tempFilePath);
                }
                catch (Exception)
                {
                    return slideAsShape;
                }
                return slideAsShape;
            }
            catch (Exception)
            {
                // It is possible that there could permissions-related issues that cause user to be unable to create/delete files.
                // In that case, we proceed with copy-pasting the slide as image.
                PowerPoint.Shape slideAsShape = targetSlide.Shapes.SafeCopySlide(slideToAdd);
                return slideAsShape;
            }
        }
    }
}
