using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.ActionFramework.Common.Log;
using PowerPointLabs.EffectsLab;
using PowerPointLabs.Utils;

using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointLabs.Models
{
    class PowerPointSpotlightSlide : PowerPointSlide
    {
#pragma warning disable 0618
        //Padding so that after cropping, circumference of spotLightPicture will not have soft edge
        private const float SoftEdgePadding = 3.0f;

        private PowerPointSpotlightSlide(PowerPoint.Slide slide) : base(slide)
        {
            if (!slide.Name.Contains("PPTLabsSpotlight"))
            {
                _slide.Name = "PPTLabsSpotlight" + DateTime.Now.ToString("yyyyMMddHHmmssffff");
            }
        }

        public static PowerPointSlide FromSlideFactory(PowerPoint.Slide slide)
        {
            if (slide == null)
            {
                return null;
            }

            return new PowerPointSpotlightSlide(slide);
        }

        public void PrepareForSpotlight()
        {
            MoveMotionAnimation();

            //Delete shapes with exit animations
            List<PowerPoint.Shape> shapes = _slide.Shapes.Cast<PowerPoint.Shape>().ToList();
            IEnumerable<PowerPoint.Shape> matchingShapes = shapes.Where(current => (HasExitAnimation(current)));
            foreach (PowerPoint.Shape s in matchingShapes)
            {
                s.SafeDelete();
            }

            foreach (PowerPoint.Shape s in _slide.Shapes)
            {
                DeleteShapeAnimations(s);
            }

            DeleteSlideNotes();
            DeleteSlideMedia();
            ManageSlideTransitions();
        }

        //Edit selected spotlight shape to fit within the current slide
        public PowerPoint.Shape CreateSpotlightShape(PowerPoint.Shape spotShape)
        {
            spotShape.Copy();
            bool isCallout = false;
            PowerPoint.Shape spotlightShape;
            
            if (spotShape.Type == Office.MsoShapeType.msoCallout)
            {
                isCallout = true;
            }

            if (isCallout)
            {
                spotlightShape = this.Shapes.Paste()[1];
                LegacyShapeUtil.CopyShapePosition(spotShape, ref spotlightShape);
            }
            else
            {
                spotlightShape = this.Shapes.PasteSpecial(PowerPoint.PpPasteDataType.ppPastePNG)[1];
                LegacyShapeUtil.CopyShapePosition(spotShape, ref spotlightShape);
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
                RenderSpotlightPicture(spotlightPicture);
                indicatorShape.ZOrder(Office.MsoZOrderCmd.msoBringToFront);
            }
            catch (Exception e)
            {
                Logger.LogException(e, "AddSpotlightEffect");
                throw;
            }
        }

        /// <summary>
        /// Export formatted spotlight picture as a new picture,
        /// then use the new pic to replace the formatted one.
        /// Thus when it's displayed, no need to render the effect (which's very slow)
        /// </summary>
        /// <param name="spotlightPicture"></param>
        private void RenderSpotlightPicture(PowerPoint.Shape spotlightPicture)
        {
            string dirOfRenderedPicture = Path.GetTempPath() + @"\rendered_" + spotlightPicture.Name;
            //Render process:
            //export formatted spotlight picture to a temp folder
            spotlightPicture.Export(dirOfRenderedPicture, PowerPoint.PpShapeFormat.ppShapeFormatPNG);
            //then add the exported new picture back
            PowerPoint.Shape renderedPicture = PowerPointCurrentPresentationInfo.CurrentSlide.Shapes.AddPicture(
                dirOfRenderedPicture, Office.MsoTriState.msoFalse, Office.MsoTriState.msoTrue,
                spotlightPicture.Left, spotlightPicture.Top, spotlightPicture.Width, spotlightPicture.Height);

            renderedPicture.Name = spotlightPicture.Name + "_rendered";

            //get rid of extra padding
            CropSpotlightPictureToSlide(ref renderedPicture);

            spotlightPicture.SafeDelete();
        }

        private void ManageSlideTransitions()
        {
            RemoveSlideTransitions();
            _slide.SlideShowTransition.AdvanceOnTime = Office.MsoTriState.msoFalse;
            _slide.SlideShowTransition.AdvanceOnClick = Office.MsoTriState.msoTrue;
        }

        private void CropSpotlightPictureToSlide(ref PowerPoint.Shape shapeToCrop)
        {
            float scaleFactorWidth = ShapeUtil.GetScaleWidth(shapeToCrop);
            float scaleFactorHeight = ShapeUtil.GetScaleHeight(shapeToCrop);

            if (shapeToCrop.Left < 0)
            {
                shapeToCrop.PictureFormat.CropLeft += ((0.0f - shapeToCrop.Left) / scaleFactorWidth);
            }
            if (shapeToCrop.Left + shapeToCrop.Width > PowerPointPresentation.Current.SlideWidth)
            {
                shapeToCrop.PictureFormat.CropRight += ((shapeToCrop.Left + shapeToCrop.Width - PowerPointPresentation.Current.SlideWidth) / scaleFactorWidth);
            }
            if (shapeToCrop.Top < 0)
            {
                shapeToCrop.PictureFormat.CropTop += ((0.0f - shapeToCrop.Top) / scaleFactorHeight);
            }
            if (shapeToCrop.Top + shapeToCrop.Height > PowerPointPresentation.Current.SlideHeight)
            {
                shapeToCrop.PictureFormat.CropBottom += ((shapeToCrop.Top + shapeToCrop.Height - PowerPointPresentation.Current.SlideHeight) / scaleFactorHeight);
            }
        }

        private void PrepareSpotlightShape(ref PowerPoint.Shape spotlightShape)
        {
            spotlightShape.Line.Visible = Office.MsoTriState.msoFalse;
            if (spotlightShape.HasTextFrame == Office.MsoTriState.msoTrue && spotlightShape.TextFrame.HasText == Office.MsoTriState.msoTrue)
            {
                spotlightShape.TextFrame.TextRange.Font.Color.RGB = 0xffffff;
            }

            spotlightShape.Name = "SpotlightShape" + Guid.NewGuid().ToString();
        }

        private void AddRectangleShape()
        {
            float softEdges = EffectsLabSettings.SpotlightSoftEdges;
            Color color = EffectsLabSettings.SpotlightColor;
            float transparency = EffectsLabSettings.SpotlightTransparency;

            PowerPoint.Shape rectangleShape = Shapes.AddShape(
                    Office.MsoAutoShapeType.msoShapeRectangle, 
                    (-SoftEdgePadding/2 * softEdges), 
                    (-SoftEdgePadding/2 * softEdges), 
                    (PowerPointPresentation.Current.SlideWidth + (SoftEdgePadding * softEdges)), 
                    (PowerPointPresentation.Current.SlideHeight + (SoftEdgePadding * softEdges)));
            rectangleShape.Fill.Solid();
            rectangleShape.Fill.ForeColor.RGB = ColorTranslator.ToWin32(color);
            rectangleShape.Fill.Transparency = transparency;
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

            // Save the original dimensions because ppPastePNG is resized in PowerPoint 2016
            float originalWidth = currentSelection.ShapeRange[1].Width;
            float originalHeight = currentSelection.ShapeRange[1].Height;
            currentSelection.Cut();

            PowerPoint.Shape spotlightPicture = this.Shapes.PasteSpecial(PowerPoint.PpPasteDataType.ppPastePNG)[1];
            spotlightPicture.Width = originalWidth;
            spotlightPicture.Height = originalHeight;

            return spotlightPicture;
        }

        private void FormatSpotlightPicture(PowerPoint.Shape spotlightPicture)
        {
            float softEdges = EffectsLabSettings.SpotlightSoftEdges;

            spotlightPicture.PictureFormat.TransparencyColor = 0xffffff;
            spotlightPicture.PictureFormat.TransparentBackground = Office.MsoTriState.msoTrue;
            spotlightPicture.Left = -SoftEdgePadding/2 * softEdges;
            spotlightPicture.Top = -SoftEdgePadding/2 * softEdges;
            spotlightPicture.LockAspectRatio = Office.MsoTriState.msoFalse;
            float incrementWidth = (SoftEdgePadding * softEdges) / spotlightPicture.Width;
            float incrementHeight = (SoftEdgePadding * softEdges) / spotlightPicture.Height;

            spotlightPicture.SoftEdge.Radius = softEdges;
            spotlightPicture.Shadow.Size = 0;

            spotlightPicture.Name = "SpotlightShape1";
        }
    }
}
