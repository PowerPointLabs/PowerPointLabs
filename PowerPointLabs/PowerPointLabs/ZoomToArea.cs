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
        public static bool multiSlideZoomChecked = true;

        public static void AddZoomToArea()
        {
            try
            {
                var currentSlide = PowerPointCurrentPresentationInfo.CurrentSlide as PowerPointSlide;
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
                PowerPointLabsGlobals.AddAckSlide();   
            }
            catch (Exception e)
            {
                PowerPointLabsGlobals.LogException(e, "AddZoomToArea");
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

        //Shape dimensions should match the slide dimensions and the shape should be within the slide
        private static PowerPoint.Shape GetBestFitShape(PowerPointSlide currentSlide, PowerPoint.Shape zoomShape)
        {
            zoomShape.Copy();
            PowerPoint.Shape zoomShapeCopy = currentSlide.Shapes.Paste()[1];

            zoomShapeCopy.LockAspectRatio = Office.MsoTriState.msoFalse;

            if (zoomShape.Width > zoomShape.Height)
            {
                zoomShapeCopy.Width = zoomShape.Width;
                zoomShapeCopy.Height = PowerPointCurrentPresentationInfo.SlideHeight * zoomShapeCopy.Width / PowerPointCurrentPresentationInfo.SlideWidth;
            }
            else
            {
                zoomShapeCopy.Height = zoomShape.Height;
                zoomShapeCopy.Width = Globals.ThisAddIn.Application.ActivePresentation.PageSetup.SlideWidth * zoomShapeCopy.Height / Globals.ThisAddIn.Application.ActivePresentation.PageSetup.SlideHeight;
            }
            PowerPointLabsGlobals.CopyShapePosition(zoomShape, ref zoomShapeCopy);

            if (zoomShapeCopy.Width > PowerPointCurrentPresentationInfo.SlideWidth)
                zoomShapeCopy.Width = PowerPointCurrentPresentationInfo.SlideWidth;
            if (zoomShapeCopy.Height > PowerPointCurrentPresentationInfo.SlideHeight)
                zoomShapeCopy.Height = PowerPointCurrentPresentationInfo.SlideHeight;

            if (zoomShapeCopy.Left < 0)
                zoomShapeCopy.Left = 0;
            if (zoomShapeCopy.Left + zoomShapeCopy.Width > PowerPointCurrentPresentationInfo.SlideWidth)
                zoomShapeCopy.Left = PowerPointCurrentPresentationInfo.SlideWidth - zoomShapeCopy.Width;
            if (zoomShapeCopy.Top < 0)
                zoomShapeCopy.Top = 0;
            if (zoomShapeCopy.Top + zoomShapeCopy.Height > PowerPointCurrentPresentationInfo.SlideHeight)
                zoomShapeCopy.Top = PowerPointCurrentPresentationInfo.SlideHeight - zoomShapeCopy.Height;

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
            if (currentSlide.Name.Contains("PPTLabsZoomToAreaSlide") && currentSlide.Index != PowerPointCurrentPresentationInfo.SlideCount)
            {
                PowerPointSlide nextSlide = PowerPointCurrentPresentationInfo.Slides.ElementAt(currentSlide.Index);
                PowerPointSlide tempSlide = null;
                while ((nextSlide.Name.Contains("PPTLabsMagnifyingSlide") || (nextSlide.Name.Contains("PPTLabsMagnifiedSlide"))
                       || (nextSlide.Name.Contains("PPTLabsDeMagnifyingSlide")) || (nextSlide.Name.Contains("PPTLabsMagnifiedPanSlide"))
                       || (nextSlide.Name.Contains("PPTLabsMagnifyingSingleSlide"))) && nextSlide.Index < PowerPointCurrentPresentationInfo.SlideCount)
                {
                    tempSlide = nextSlide;
                    nextSlide = PowerPointCurrentPresentationInfo.Slides.ElementAt(tempSlide.Index);
                    tempSlide.Delete();
                }
            }
        }
    }
}
