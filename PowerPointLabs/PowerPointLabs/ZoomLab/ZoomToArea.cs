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
    internal static class ZoomToArea
    {
#pragma warning disable 0618
        public static void AddZoomToArea()
        {
            if (!IsSelectingShapes())
            {
                return;
            }

            try
            {
                PowerPointSlide currentSlide = PowerPointCurrentPresentationInfo.CurrentSlide;
                DeleteExistingZoomToAreaSlides(currentSlide);
                currentSlide.Name = "PPTLabsZoomToAreaSlide" + DateTime.Now.ToString("yyyyMMddHHmmssffff");

                PowerPoint.ShapeRange selectedShapes = Globals.ThisAddIn.Application.ActiveWindow.Selection.ShapeRange;
                List<PowerPoint.Shape> zoomRectangles = ReplaceWithZoomRectangleImages(currentSlide, selectedShapes);

                MakeInvisible(zoomRectangles);
                List<PowerPoint.Shape> editedSelectedShapes = GetEditedShapesForZoomToArea(currentSlide, zoomRectangles);

                List<PowerPointSlide> addedSlides = AddMultiSlideZoomToArea(currentSlide, editedSelectedShapes);
                if (!ZoomLabSettings.MultiSlideZoomChecked)
                {
                    SlideUtil.SquashSlides(addedSlides);
                }

                MakeVisible(zoomRectangles);

                Globals.ThisAddIn.Application.ActiveWindow.View.GotoSlide(currentSlide.Index);
                PowerPointPresentation.Current.AddAckSlide();

                // Always call ReleaseComObject and GC.Collect after shape deletion to prevent shape corruption after undo.
                System.Runtime.InteropServices.Marshal.ReleaseComObject(selectedShapes);
                GC.Collect();
            }
            catch (Exception e)
            {
                Logger.LogException(e, "AddZoomToArea");
                ErrorDialogBox.ShowDialog("Error when adding zoom to area", "An error occurred when adding zoom to area.", e);
                throw;
            }
        }

        private static List<PowerPointSlide> AddMultiSlideZoomToArea(PowerPointSlide currentSlide, List<PowerPoint.Shape> shapesToZoom)
        {
            List<PowerPointSlide> addedSlides = new List<PowerPointSlide>();

            int shapeCount = 1;
            PowerPointSlide lastMagnifiedSlide = null;
            PowerPointMagnifyingSlide magnifyingSlide = null;
            PowerPointMagnifiedSlide magnifiedSlide = null;
            PowerPointMagnifiedPanSlide magnifiedPanSlide = null;
            PowerPointDeMagnifyingSlide deMagnifyingSlide = null;

            foreach (PowerPoint.Shape selectedShape in shapesToZoom)
            {
                magnifyingSlide = (PowerPointMagnifyingSlide)currentSlide.CreateZoomMagnifyingSlide();
                magnifyingSlide.AddZoomToAreaAnimation(selectedShape);

                magnifiedSlide = (PowerPointMagnifiedSlide)magnifyingSlide.CreateZoomMagnifiedSlide();
                magnifiedSlide.AddZoomToAreaAnimation(selectedShape);
                addedSlides.Add(magnifiedSlide);

                if (shapeCount != 1)
                {
                    magnifiedPanSlide = (PowerPointMagnifiedPanSlide)lastMagnifiedSlide.CreateZoomPanSlide();
                    magnifiedPanSlide.AddZoomToAreaAnimation(lastMagnifiedSlide, magnifiedSlide);
                    addedSlides.Add(magnifiedPanSlide);
                }

                if (shapeCount == shapesToZoom.Count)
                {
                    deMagnifyingSlide = (PowerPointDeMagnifyingSlide)magnifyingSlide.CreateZoomDeMagnifyingSlide();
                    deMagnifyingSlide.MoveTo(magnifyingSlide.Index + 2);
                    deMagnifyingSlide.AddZoomToAreaAnimation(selectedShape);
                    addedSlides.Add(deMagnifyingSlide);
                }

                selectedShape.SafeDelete();

                if (shapeCount != 1)
                {
                    magnifyingSlide.Delete();
                    magnifiedSlide.MoveTo(magnifiedPanSlide.Index);
                    if (deMagnifyingSlide != null)
                    {
                        deMagnifyingSlide.MoveTo(magnifiedSlide.Index);
                    }

                    lastMagnifiedSlide = magnifiedSlide;
                }
                else
                {
                    addedSlides.Add(magnifyingSlide);
                    lastMagnifiedSlide = magnifiedSlide;
                }

                shapeCount++;
            }

            SlideUtil.SortByIndex(addedSlides);
            return addedSlides;
        }

        private static List<PowerPoint.Shape> ReplaceWithZoomRectangleImages(PowerPointSlide currentSlide, PowerPoint.ShapeRange shapeRange)
        {
            List<PowerPoint.Shape> zoomRectangles = new List<PowerPoint.Shape>();
            int shapeCount = 1;
            foreach (PowerPoint.Shape zoomShape in shapeRange)
            {
                PowerPoint.Shape zoomRectangle = currentSlide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeRectangle,
                                                                zoomShape.Left,
                                                                zoomShape.Top,
                                                                zoomShape.Width,
                                                                zoomShape.Height);
                currentSlide.AddAppearDisappearAnimation(zoomRectangle);

                // Set Name
                zoomRectangle.Name = "PPTLabsMagnifyShape" + DateTime.Now.ToString("yyyyMMddHHmmssffff");

                // Set Text
                zoomRectangle.TextFrame2.TextRange.Text = "Zoom Shape " + shapeCount;
                zoomRectangle.TextFrame2.AutoSize = Office.MsoAutoSize.msoAutoSizeTextToFitShape;
                zoomRectangle.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = 0xffffff;
                zoomRectangle.TextFrame2.TextRange.Font.Bold = Office.MsoTriState.msoTrue;

                // Set Color
                zoomRectangle.Fill.ForeColor.RGB = 0xaaaaaa;
                zoomRectangle.Fill.Transparency = 0.7f;
                zoomRectangle.Line.ForeColor.RGB = 0x000000;

                zoomRectangles.Add(zoomRectangle);
                zoomShape.SafeDelete();
                shapeCount++;
            }
            return zoomRectangles;
        }

        private static List<PowerPoint.Shape> GetEditedShapesForZoomToArea(PowerPointSlide currentSlide, List<PowerPoint.Shape> zoomRectangles)
        {
            return zoomRectangles.Select(zoomShape => GetBestFitShape(currentSlide, zoomShape)).ToList();
        }

        //Shape dimensions should match the slide dimensions and the shape should be within the slide
        private static PowerPoint.Shape GetBestFitShape(PowerPointSlide currentSlide, PowerPoint.Shape zoomShape)
        {
            PowerPoint.Shape zoomShapeCopy = zoomShape.Duplicate()[1];
            
            zoomShapeCopy.LockAspectRatio = Office.MsoTriState.msoFalse;

            if (zoomShape.Width > zoomShape.Height)
            {
                zoomShapeCopy.Width = zoomShape.Width;
                zoomShapeCopy.Height = PowerPointPresentation.Current.SlideHeight * zoomShapeCopy.Width / PowerPointPresentation.Current.SlideWidth;
            }
            else
            {
                zoomShapeCopy.Height = zoomShape.Height;
                zoomShapeCopy.Width = PowerPointPresentation.Current.SlideWidth * zoomShapeCopy.Height / PowerPointPresentation.Current.SlideHeight;
            }
            LegacyShapeUtil.CopyCenterShapePosition(zoomShape, ref zoomShapeCopy);

            if (zoomShapeCopy.Width > PowerPointPresentation.Current.SlideWidth)
            {
                zoomShapeCopy.Width = PowerPointPresentation.Current.SlideWidth;
            }

            if (zoomShapeCopy.Height > PowerPointPresentation.Current.SlideHeight)
            {
                zoomShapeCopy.Height = PowerPointPresentation.Current.SlideHeight;
            }

            if (zoomShapeCopy.Left < 0)
            {
                zoomShapeCopy.Left = 0;
            }

            if (zoomShapeCopy.Left + zoomShapeCopy.Width > PowerPointPresentation.Current.SlideWidth)
            {
                zoomShapeCopy.Left = PowerPointPresentation.Current.SlideWidth - zoomShapeCopy.Width;
            }

            if (zoomShapeCopy.Top < 0)
            {
                zoomShapeCopy.Top = 0;
            }

            if (zoomShapeCopy.Top + zoomShapeCopy.Height > PowerPointPresentation.Current.SlideHeight)
            {
                zoomShapeCopy.Top = PowerPointPresentation.Current.SlideHeight - zoomShapeCopy.Height;
            }

            return zoomShapeCopy;
        }

        private static void MakeInvisible(IEnumerable<PowerPoint.Shape> zoomRectangles)
        {
            foreach (PowerPoint.Shape sh in zoomRectangles)
            {
                sh.Visible = Office.MsoTriState.msoFalse;
            }
        }

        private static void MakeVisible(IEnumerable<PowerPoint.Shape> zoomRectangles)
        {
            foreach (PowerPoint.Shape sh in zoomRectangles)
            {
                sh.Visible = Office.MsoTriState.msoTrue;
            }
        }

        private static bool IsSelectingShapes()
        {
            PowerPoint.Selection selection = Globals.ThisAddIn.Application.ActiveWindow.Selection;
            return selection.Type == PowerPoint.PpSelectionType.ppSelectionShapes && selection.ShapeRange.Count > 0;
        }

        private static void DeleteExistingZoomToAreaSlides(PowerPointSlide currentSlide)
        {
            if (currentSlide.Name.Contains("PPTLabsZoomToAreaSlide") && currentSlide.Index != PowerPointPresentation.Current.SlideCount)
            {
                PowerPointSlide nextSlide = PowerPointPresentation.Current.Slides[currentSlide.Index];
                while ((nextSlide.Name.Contains("PPTLabsMagnifyingSlide") || (nextSlide.Name.Contains("PPTLabsMagnifiedSlide"))
                       || (nextSlide.Name.Contains("PPTLabsDeMagnifyingSlide")) || (nextSlide.Name.Contains("PPTLabsMagnifiedPanSlide"))
                       || (nextSlide.Name.Contains("PPTLabsMagnifyingSingleSlide"))) && nextSlide.Index < PowerPointPresentation.Current.SlideCount)
                {
                    PowerPointSlide tempSlide = nextSlide;
                    nextSlide = PowerPointPresentation.Current.Slides[tempSlide.Index];
                    tempSlide.Delete();
                }
            }
        }
    }
}
