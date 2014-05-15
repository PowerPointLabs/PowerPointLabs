using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Runtime.InteropServices;
using PowerPointLabs.Models;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointLabs
{
    class ZoomToArea
    {
        public static bool backgroundZoomChecked = true;
        public static bool multiSlideZoomChecked = false;

        public static void AddZoomToArea()
        {
            try
            {
                var currentSlide = PowerPointPresentation.CurrentSlide as PowerPointSlide;
                DeleteExistingZoomToAreaSlides(currentSlide);
                currentSlide.Name = "PPTLabsZoomToAreaSlide" + DateTime.Now.ToString("yyyyMMddHHmmssffff");

                PowerPoint.ShapeRange selectedShapes = Globals.ThisAddIn.Application.ActiveWindow.Selection.ShapeRange;
                List<PowerPoint.Shape> editedSelectedShapes = GetEditedShapesForZoomToArea(currentSlide, selectedShapes);

                if (!multiSlideZoomChecked)
                    AddSingleSlideZoomToArea(currentSlide, editedSelectedShapes);
                else
                    AddMultiSlideZoomToArea(currentSlide, editedSelectedShapes);

                Globals.ThisAddIn.Application.ActiveWindow.View.GotoSlide(currentSlide.Index);
                PostFormatSelectedShapes(ref selectedShapes);
                AddAckSlide();   
            }
            catch (Exception e)
            {
                //LogException(e, "AddDrillDownAnimation");
                throw;
            }
        }

        private static void AddMultiSlideZoomToArea(PowerPointSlide currentSlide, List<PowerPoint.Shape> shapesToZoom)
        {
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

                if (shapeCount != 1)
                {
                    magnifiedPanSlide = (PowerPointMagnifiedPanSlide)lastMagnifiedSlide.CreateZoomPanSlide();
                    magnifiedPanSlide.AddZoomToAreaAnimation(lastMagnifiedSlide, magnifiedSlide);
                }

                if (shapeCount == shapesToZoom.Count)
                {
                    deMagnifyingSlide = (PowerPointDeMagnifyingSlide)magnifyingSlide.CreateZoomDeMagnifyingSlide();
                    deMagnifyingSlide.MoveTo(magnifyingSlide.Index + 2);
                    deMagnifyingSlide.AddZoomToAreaAnimation(selectedShape);
                }

                selectedShape.Delete();

                if (shapeCount != 1)
                {
                    magnifyingSlide.Delete();
                    magnifiedSlide.MoveTo(magnifiedPanSlide.Index);
                    if (deMagnifyingSlide != null)
                        deMagnifyingSlide.MoveTo(magnifiedSlide.Index);
                    lastMagnifiedSlide = magnifiedSlide;
                }
                else
                {
                    lastMagnifiedSlide = magnifiedSlide;
                }

                shapeCount++;
            }
        }

        private static void AddSingleSlideZoomToArea(PowerPointSlide currentSlide, List<PowerPoint.Shape> shapesToZoom)
        {
            var zoomSlide = currentSlide.CreateZoomToAreaSingleSlide() as PowerPointZoomToAreaSingleSlide;
            zoomSlide.PrepareForZoomToArea(shapesToZoom);
            zoomSlide.AddZoomToAreaAnimation(currentSlide, shapesToZoom);
        }

        private static List<PowerPoint.Shape> GetEditedShapesForZoomToArea(PowerPointSlide currentSlide, PowerPoint.ShapeRange selectedShapes)
        {
            List<PowerPoint.Shape> editedSelectedShapes = new List<PowerPoint.Shape>();
            int shapeCount = 1;
            foreach (PowerPoint.Shape zoomShape in selectedShapes)
            {
                currentSlide.DeleteShapeAnimations(zoomShape);
                currentSlide.AddAppearDisappearAnimation(zoomShape);
                zoomShape.Name = "PPTLabsMagnifyShape" + DateTime.Now.ToString("yyyyMMddHHmmssffff");
                editedSelectedShapes.Add(GetBestFitShape(currentSlide, zoomShape));

                if (zoomShape.HasTextFrame == Office.MsoTriState.msoTrue)
                {
                    zoomShape.TextFrame2.DeleteText();
                    zoomShape.TextFrame2.TextRange.Text = "Zoom Shape " + shapeCount;
                    zoomShape.TextFrame2.AutoSize = Office.MsoAutoSize.msoAutoSizeTextToFitShape;
                    zoomShape.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = 0xffffff;
                    zoomShape.TextFrame2.TextRange.Font.Bold = Office.MsoTriState.msoTrue;
                }

                zoomShape.Visible = Office.MsoTriState.msoFalse;
                shapeCount++;
            }
            return editedSelectedShapes;
        }

        private static PowerPoint.Shape GetBestFitShape(PowerPointSlide currentSlide, PowerPoint.Shape zoomShape)
        {
            zoomShape.Copy();
            PowerPoint.Shape zoomShapeCopy = currentSlide.Shapes.Paste()[1];

            zoomShapeCopy.LockAspectRatio = Office.MsoTriState.msoFalse;

            if (zoomShape.Width > zoomShape.Height)
            {
                zoomShapeCopy.Width = zoomShape.Width;
                zoomShapeCopy.Height = PowerPointPresentation.SlideHeight * zoomShapeCopy.Width / PowerPointPresentation.SlideWidth;
            }
            else
            {
                zoomShapeCopy.Height = zoomShape.Height;
                zoomShapeCopy.Width = Globals.ThisAddIn.Application.ActivePresentation.PageSetup.SlideWidth * zoomShapeCopy.Height / Globals.ThisAddIn.Application.ActivePresentation.PageSetup.SlideHeight;
            }
            CopyShapePosition(zoomShape, ref zoomShapeCopy);

            if (zoomShapeCopy.Width > PowerPointPresentation.SlideWidth)
                zoomShapeCopy.Width = PowerPointPresentation.SlideWidth;
            if (zoomShapeCopy.Height > PowerPointPresentation.SlideHeight)
                zoomShapeCopy.Height = PowerPointPresentation.SlideHeight;

            if (zoomShapeCopy.Left < 0)
                zoomShapeCopy.Left = 0;
            if (zoomShapeCopy.Left + zoomShapeCopy.Width > PowerPointPresentation.SlideWidth)
                zoomShapeCopy.Left = PowerPointPresentation.SlideWidth - zoomShapeCopy.Width;
            if (zoomShapeCopy.Top < 0)
                zoomShapeCopy.Top = 0;
            if (zoomShapeCopy.Top + zoomShapeCopy.Height > PowerPointPresentation.SlideHeight)
                zoomShapeCopy.Top = PowerPointPresentation.SlideHeight - zoomShapeCopy.Height;

            return zoomShapeCopy;
        }

        private static void PostFormatSelectedShapes(ref PowerPoint.ShapeRange selectedShapes)
        {
            foreach (PowerPoint.Shape sh in selectedShapes)
            {
                sh.Visible = Office.MsoTriState.msoTrue;
                sh.Fill.ForeColor.RGB = 0xaaaaaa;
                sh.Fill.Transparency = 0.7f;
                sh.Line.ForeColor.RGB = 0x000000;
            }
        }

        private static void DeleteExistingZoomToAreaSlides(PowerPointSlide currentSlide)
        {
            if (currentSlide.Name.Contains("PPTLabsZoomToAreaSlide") && currentSlide.Index != PowerPointPresentation.SlideCount)
            {
                PowerPointSlide nextSlide = PowerPointPresentation.Slides.ElementAt(currentSlide.Index);
                PowerPointSlide tempSlide = null;
                while ((nextSlide.Name.Contains("PPTLabsMagnifyingSlide") || (nextSlide.Name.Contains("PPTLabsMagnifiedSlide"))
                       || (nextSlide.Name.Contains("PPTLabsDeMagnifyingSlide")) || (nextSlide.Name.Contains("PPTLabsMagnifiedPanSlide")))
                       && nextSlide.Index < PowerPointPresentation.SlideCount)
                {
                    tempSlide = nextSlide;
                    nextSlide = PowerPointPresentation.Slides.ElementAt(tempSlide.Index);
                    tempSlide.Delete();
                }
            }
        }

        private static void CopyShapePosition(PowerPoint.Shape shapeToCopy, ref PowerPoint.Shape shapeToMove)
        {
            shapeToMove.Left = shapeToCopy.Left + (shapeToCopy.Width / 2) - (shapeToMove.Width / 2);
            shapeToMove.Top = shapeToCopy.Top + (shapeToCopy.Height / 2) - (shapeToMove.Height / 2);
        }

        private static void CopyShapeSize(PowerPoint.Shape shapeToCopy, ref PowerPoint.Shape shapeToMove)
        {
            shapeToMove.LockAspectRatio = Office.MsoTriState.msoFalse;
            shapeToMove.Width = shapeToCopy.Width;
            shapeToMove.Height = shapeToCopy.Height;
        }

        private static void CopyShapeAttributes(PowerPoint.Shape shapeToCopy, ref PowerPoint.Shape shapeToMove)
        {
            CopyShapeSize(shapeToCopy, ref shapeToMove);
            CopyShapePosition(shapeToCopy, ref shapeToMove);
        }

        private static void FitShapeToSlide(ref PowerPoint.Shape shapeToMove)
        {
            shapeToMove.LockAspectRatio = Office.MsoTriState.msoFalse;
            shapeToMove.Left = 0;
            shapeToMove.Top = 0;
            shapeToMove.Width = PowerPointPresentation.SlideWidth;
            shapeToMove.Height = PowerPointPresentation.SlideHeight;
        }

        private static void AddAckSlide()
        {
            try
            {
                PowerPointSlide lastSlide = PowerPointPresentation.Slides.Last();
                if (!lastSlide.isAckSlide())
                {
                    lastSlide.CreateAckSlide();
                }
            }
            catch (Exception e)
            {
                //LogException(e, "AddAckSlide");
                throw;
            }
        }
    }
}
