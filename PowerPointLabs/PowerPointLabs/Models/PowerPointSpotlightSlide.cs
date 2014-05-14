using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointLabs.Models
{
    class PowerPointSpotlightSlide : PowerPointSlide
    {
        private PowerPointSpotlightSlide(PowerPoint.Slide slide) : base(slide)
        {
            if (!slide.Name.Contains("PPTLabsSpotlight"))
                _slide.Name = "PPTLabsSpotlight" + DateTime.Now.ToString("yyyyMMddHHmmssffff");
        }

        new public static PowerPointSlide FromSlideFactory(PowerPoint.Slide slide)
        {
            if (slide == null)
            {
                return null;
            }

            return new PowerPointSpotlightSlide(slide);
        }

        public void PrepareForSpotlight()
        {
            DeleteSlideNotes();
            DeleteSlideMedia();
            ManageSlideTransitions();
        }

        public PowerPoint.Shape CreateSpotlightShape(PowerPoint.Shape spotShape)
        {
            spotShape.Copy();
            bool isCallout = false;
            PowerPoint.Shape spotlightShape;
            if (spotShape.Type == Office.MsoShapeType.msoCallout)
                isCallout = true;

            if (isCallout)
            {
                spotlightShape = this.Shapes.Paste()[1];
                MoveToOriginalPosition(spotShape, ref spotlightShape);
            }
            else
            {
                spotlightShape = this.Shapes.PasteSpecial(PowerPoint.PpPasteDataType.ppPastePNG)[1];
                MoveToOriginalPosition(spotShape, ref spotlightShape);
                CropSpotlightPictureToSlide(ref spotlightShape);
            }

            PrepareSpotlightShape(ref spotlightShape);
            return spotlightShape;
        }

        public void AddSpotlightEffect(List<PowerPoint.Shape> spotlightShapes)
        {
            try
            {
                PowerPoint.Shape indicatorShape = AddPowerPointLabsIndicator();
                AddRectangleShape();
                PowerPoint.Shape spotlightPicture = ConvertToSpotlightPicture(spotlightShapes);
                FormatSpotlightPicture(spotlightPicture);
                indicatorShape.ZOrder(Office.MsoZOrderCmd.msoBringToFront);
            }
            catch (Exception e)
            {
                //LogException(e, "AddSpotlightEffect");
                throw;
            }
        }

        private void ManageSlideTransitions()
        {
            base.RemoveSlideTransitions();
            _slide.SlideShowTransition.AdvanceOnTime = Office.MsoTriState.msoFalse;
            _slide.SlideShowTransition.AdvanceOnClick = Office.MsoTriState.msoTrue;
        }

        private void MoveToOriginalPosition(PowerPoint.Shape reference, ref PowerPoint.Shape spotlightShape)
        {
            spotlightShape.Left = reference.Left + (reference.Width / 2) - (spotlightShape.Width / 2);
            spotlightShape.Top = reference.Top + (reference.Height / 2) - (spotlightShape.Height / 2);
        }

        private void CropSpotlightPictureToSlide(ref PowerPoint.Shape shapeToCrop)
        {
            if (shapeToCrop.Left < 0)
            {
                shapeToCrop.PictureFormat.CropLeft += (0.0f - shapeToCrop.Left);
            }
            if (shapeToCrop.Left + shapeToCrop.Width > PowerPointPresentation.SlideWidth)
            {
                shapeToCrop.PictureFormat.CropRight += (shapeToCrop.Left + shapeToCrop.Width - PowerPointPresentation.SlideWidth);
            }
            if (shapeToCrop.Top < 0)
            {
                shapeToCrop.PictureFormat.CropTop += (0.0f - shapeToCrop.Top);
            }
            if (shapeToCrop.Top + shapeToCrop.Height > PowerPointPresentation.SlideHeight)
            {
                shapeToCrop.PictureFormat.CropBottom += (shapeToCrop.Top + shapeToCrop.Height - PowerPointPresentation.SlideHeight);
            }
        }

        private void PrepareSpotlightShape(ref PowerPoint.Shape spotlightShape)
        {
            spotlightShape.Line.Visible = Office.MsoTriState.msoFalse;
            if (spotlightShape.HasTextFrame == Office.MsoTriState.msoTrue && spotlightShape.TextFrame.HasText == Office.MsoTriState.msoTrue)
                spotlightShape.TextFrame.TextRange.Font.Color.RGB = 0xffffff;
            spotlightShape.Name = "SpotlightShape" + Guid.NewGuid().ToString();
        }

        private void AddRectangleShape()
        {
            PowerPoint.Shape rectangleShape = this.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeRectangle, (-1 * Spotlight.defaultSoftEdges), (-1 * Spotlight.defaultSoftEdges), (PowerPointPresentation.SlideWidth + (2.0f * Spotlight.defaultSoftEdges)), (PowerPointPresentation.SlideHeight + (2.0f * Spotlight.defaultSoftEdges)));
            rectangleShape.Fill.ForeColor.RGB = 0x000000;
            rectangleShape.Fill.Transparency = Spotlight.defaultTransparency;
            rectangleShape.Line.Visible = Office.MsoTriState.msoFalse;
            rectangleShape.Name = "SpotlightShape1";
            rectangleShape.ZOrder(Office.MsoZOrderCmd.msoSendToBack);
        }

        private PowerPoint.Shape ConvertToSpotlightPicture(List<PowerPoint.Shape> spotlightShapes)
        {
            Globals.ThisAddIn.Application.ActiveWindow.View.GotoSlide(this.Index);
            List<String> shapeNames = new List<String>();
            shapeNames.Add("SpotlightShape1");
            foreach (PowerPoint.Shape sh in spotlightShapes)
            {
                shapeNames.Add(sh.Name);
            }
            String[] shapeNamesArray = shapeNames.ToArray();
            PowerPoint.ShapeRange newRange = this.Shapes.Range(shapeNamesArray);
            newRange.Select();

            PowerPoint.Selection currentSelection = Globals.ThisAddIn.Application.ActiveWindow.Selection;
            currentSelection.Cut();

            PowerPoint.Shape spotlightPicture = this.Shapes.PasteSpecial(PowerPoint.PpPasteDataType.ppPastePNG)[1];
            return spotlightPicture;
        }

        private void FormatSpotlightPicture(PowerPoint.Shape spotlightPicture)
        {
            spotlightPicture.PictureFormat.TransparencyColor = 0xffffff;
            spotlightPicture.PictureFormat.TransparentBackground = Office.MsoTriState.msoTrue;
            spotlightPicture.Left = -1 * Spotlight.defaultSoftEdges;
            spotlightPicture.Top = -1 * Spotlight.defaultSoftEdges;
            spotlightPicture.LockAspectRatio = Office.MsoTriState.msoFalse;
            float incrementWidth = (2.0f * Spotlight.defaultSoftEdges) / spotlightPicture.Width;
            float incrementHeight = (2.0f * Spotlight.defaultSoftEdges) / spotlightPicture.Height;

            spotlightPicture.SoftEdge.Radius = Spotlight.defaultSoftEdges;
            spotlightPicture.Name = "SpotlightShape1";
        }
    }
}
