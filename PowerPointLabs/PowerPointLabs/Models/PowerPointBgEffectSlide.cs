using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using ImageProcessor;
using ImageProcessor.Imaging.Filters;
using Core = Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;

namespace PowerPointLabs.Models
{
    class PowerPointBgEffectSlide : PowerPointSlide
    {
        private static readonly string AnimatedBackgroundPath = Path.Combine(Path.GetTempPath(), "animatedSlide.png");

        # region Constructor
        private PowerPointBgEffectSlide(Slide slide) : base(slide)
        {
            AddPowerPointLabsIndicator().ZOrder(Core.MsoZOrderCmd.msoBringToFront);
        }

        public new static PowerPointSlide FromSlideFactory(Slide refSlide)
        {
            if (refSlide == null)
            {
                return null;
            }

            // here we cut-paste the shape to get a reference of those shapes
            var oriShapeRange = refSlide.Shapes.Paste();

            if (!CropToShape.VerifyIsShapeRangeValid(oriShapeRange))
            {
                return null;
            }

            // TODO: make use of PowerPointLabs.Presentation Model!!!
            // add new blank slide into current slides collection
            var curPresentation = PowerPointCurrentPresentationInfo.CurrentPresentation;
            var curSlideIndex = PowerPointCurrentPresentationInfo.CurrentSlide.Index;
            var customLayout = curPresentation.SlideMaster.CustomLayouts[PpSlideLayout.ppLayoutText];
            var rawSlide = curPresentation.Slides.AddSlide(curSlideIndex + 1, customLayout);
            var newSlide = PowerPointSlide.FromSlideFactory(rawSlide);

            newSlide.DeleteAllShapes();

            // get a copy of original cover shapes
            var copyShapeRange = newSlide.Shapes.Paste();
            
            // make the range invisible before animated the slide
            oriShapeRange.Visible = Core.MsoTriState.msoFalse;
            copyShapeRange.Visible = Core.MsoTriState.msoFalse;

            MakeAnimatedBackground(newSlide, refSlide);

            copyShapeRange.Visible = Core.MsoTriState.msoCTrue;
            oriShapeRange.Visible = Core.MsoTriState.msoCTrue;
            
            // crop in the original slide and put into clipboard
            var croppedShape = MakeFrontImage(oriShapeRange);

            if (croppedShape == null) return null;

            croppedShape.Cut();

            // swap the uncropped shapes and cropped shapes
            var pastedCrop = newSlide.Shapes.Paste();
            
            // calibrate pasted shapes
            pastedCrop.Left -= 12;
            pastedCrop.Top -= 12;

            copyShapeRange.Cut();
            refSlide.Shapes.Paste().Select();

            return new PowerPointBgEffectSlide(rawSlide);
        }
        # endregion

        # region API
        public void BlurBackground()
        {
            AddBackgroundImage(null);
        }

        public void GreyScaleBackground()
        {
            AddBackgroundImage(MatrixFilters.GreyScale);
        }

        public void BlackWhiteBackground()
        {
            AddBackgroundImage(MatrixFilters.BlackWhite);
        }

        public void SepiaBackground()
        {
            AddBackgroundImage(MatrixFilters.Sepia);
        }

        public void GothamBackground()
        {
            AddBackgroundImage(MatrixFilters.Gotham);
        }
        # endregion

        # region Helper Functions
        private void AddBackgroundImage(IMatrixFilter filter)
        {
            using (var imageFactory = new ImageFactory())
            {
                var image = imageFactory.Load(AnimatedBackgroundPath);

                image = filter == null ? image.GaussianBlur(20) : image.Filter(filter);

                image.Save(AnimatedBackgroundPath);
            }

            var newBackground = Shapes.AddPicture(AnimatedBackgroundPath, Core.MsoTriState.msoFalse,
                                                  Core.MsoTriState.msoTrue,
                                                  0, 0,
                                                  PowerPointCurrentPresentationInfo.SlideWidth,
                                                  PowerPointCurrentPresentationInfo.SlideHeight);

            newBackground.ZOrder(Core.MsoZOrderCmd.msoSendToBack);
        }

        private static Shape MakeFrontImage(ShapeRange shapeRange)
        {
            // soften cropped shape's edge
            shapeRange.SoftEdge.Type = Core.MsoSoftEdgeType.msoSoftEdgeType5;

            return CropToShape.Crop(shapeRange);
        }

        private static void MakeAnimatedBackground(PowerPointSlide curSlide, Slide refSlide)
        {
            // copy all shapes from ref slide to current slide
            refSlide.Shapes.Range().Copy();
            var copiedShapes = curSlide.Shapes.Paste();

            foreach (var shape in copiedShapes.Cast<Shape>().Where(curSlide.HasExitAnimation))
            {
                shape.Delete();
            }

            curSlide.MoveMotionAnimation();

            Utils.Graphics.ExportSlide(curSlide, AnimatedBackgroundPath);
            
            copiedShapes.Delete();
        }
        # endregion
    }
}
