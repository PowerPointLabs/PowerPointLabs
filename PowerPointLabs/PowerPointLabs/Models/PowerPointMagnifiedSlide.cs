using System;
using System.Collections.Generic;
using System.Linq;
using PowerPointLabs.ActionFramework.Common.Extension;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointLabs.Models
{
    class PowerPointMagnifiedSlide : PowerPointSlide
    {
#pragma warning disable 0618
        private PowerPoint.Shape indicatorShape = null;
        private PowerPoint.Shape zoomSlideCroppedShapes = null;

        private PowerPointMagnifiedSlide(PowerPoint.Slide slide) : base(slide)
        {
            _slide.Name = "PPTLabsMagnifiedSlide" + DateTime.Now.ToString("yyyyMMddHHmmssffff");
        }

        public static PowerPointSlide FromSlideFactory(PowerPoint.Slide slide)
        {
            if (slide == null)
            {
                return null;
            }

            return new PowerPointMagnifiedSlide(slide);
        }

        public void AddZoomToAreaAnimation(PowerPoint.Shape zoomShape)
        {
            PrepareForZoomToArea();

            //Create zoomed-in version of the part of the slide specified by zoom shape
            zoomSlideCroppedShapes.PictureFormat.CropLeft += zoomShape.Left;
            zoomSlideCroppedShapes.PictureFormat.CropTop += zoomShape.Top;
            zoomSlideCroppedShapes.PictureFormat.CropRight += (PowerPointPresentation.Current.SlideWidth - (zoomShape.Left + zoomShape.Width));
            zoomSlideCroppedShapes.PictureFormat.CropBottom += (PowerPointPresentation.Current.SlideHeight - (zoomShape.Top + zoomShape.Height));

            LegacyShapeUtil.CopyCenterShapePosition(zoomShape, ref zoomSlideCroppedShapes);

            zoomSlideCroppedShapes.LockAspectRatio = Office.MsoTriState.msoTrue;
            if (zoomSlideCroppedShapes.Width > zoomSlideCroppedShapes.Height)
            {
                zoomSlideCroppedShapes.Width = PowerPointPresentation.Current.SlideWidth;
            }
            else
            {
                zoomSlideCroppedShapes.Height = PowerPointPresentation.Current.SlideHeight;
            }

            zoomSlideCroppedShapes.Left = (PowerPointPresentation.Current.SlideWidth / 2) - (zoomSlideCroppedShapes.Width / 2);
            zoomSlideCroppedShapes.Top = (PowerPointPresentation.Current.SlideHeight / 2) - (zoomSlideCroppedShapes.Height / 2);

            zoomSlideCroppedShapes.PictureFormat.CropLeft = 0;
            zoomSlideCroppedShapes.PictureFormat.CropTop = 0;
            zoomSlideCroppedShapes.PictureFormat.CropRight = 0;
            zoomSlideCroppedShapes.PictureFormat.CropBottom = 0;

            indicatorShape.ZOrder(Office.MsoZOrderCmd.msoBringToFront);
        }

        private void PrepareForZoomToArea()
        {
            //Delete all shapes on slide except slide-size crop copied from magnifying slide
            List<PowerPoint.Shape> shapes = _slide.Shapes.Cast<PowerPoint.Shape>().ToList();
            IEnumerable<PowerPoint.Shape> matchingShapes = shapes.Where(current => (!current.Name.Contains("PPTLabsMagnifyAreaGroup")));
            foreach (PowerPoint.Shape s in matchingShapes)
            {
                s.SafeDelete();
            }

            zoomSlideCroppedShapes = GetShapesWithPrefix("PPTLabsMagnifyAreaGroup")[0];
            zoomSlideCroppedShapes.Visible = Office.MsoTriState.msoTrue;
            DeleteShapeAnimations(zoomSlideCroppedShapes);

            DeleteSlideNotes();
            DeleteSlideMedia();
            ManageSlideTransitions();
            indicatorShape = AddPowerPointLabsIndicator();
        }

        private void ManageSlideTransitions()
        {
            RemoveSlideTransitions();
            _slide.SlideShowTransition.AdvanceOnTime = Office.MsoTriState.msoFalse;
            _slide.SlideShowTransition.AdvanceOnClick = Office.MsoTriState.msoTrue;
        }
    }
}
