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
            // cut the original shape cover again and duplicate the slide
            // here the slide will be duplicated without the original shape cover
            oriShapeRange.Cut();
            var newSlide = PowerPointSlide.FromSlideFactory(refSlide.Duplicate()[1]);
            
            // get a copy of original cover shapes
            var copyShapeRange = newSlide.Shapes.Paste();
            // paste the original shape cover back
            oriShapeRange = refSlide.Shapes.Paste();
            
            // make the range invisible before animated the slide
            copyShapeRange.Visible = Core.MsoTriState.msoFalse;

            MakeAnimatedBackground(newSlide);

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

            return new PowerPointBgEffectSlide(newSlide.GetNativeSlide());
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

        private static void MakeAnimatedBackground(PowerPointSlide curSlide)
        {
            foreach (var shape in curSlide.Shapes.Cast<Shape>().Where(curSlide.HasExitAnimation))
            {
                shape.Delete();
            }

            curSlide.MoveMotionAnimation();

            Utils.Graphics.ExportSlide(curSlide, AnimatedBackgroundPath);

            var visibleShape = curSlide.Shapes.Cast<Shape>().Where(x => x.Visible == Core.MsoTriState.msoTrue).ToList();
            
            foreach (var shape in visibleShape)
            {
                shape.Delete();
            }
        }
        # endregion
    }
}
