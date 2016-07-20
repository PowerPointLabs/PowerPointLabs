using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using ImageProcessor;
using ImageProcessor.Imaging.Filters;
using Core = Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;

namespace PowerPointLabs.Models
{
    class PowerPointBgEffectSlide : PowerPointSlide
    {
#pragma warning disable 0618
        private static readonly string AnimatedBackgroundPath = Path.Combine(Path.GetTempPath(), "animatedSlide.png");

        # region Constructor
        private PowerPointBgEffectSlide(Slide slide) : base(slide)
        {
            AddPowerPointLabsIndicator().ZOrder(Core.MsoZOrderCmd.msoBringToFront);
        }

        public static PowerPointBgEffectSlide BgEffectFactory(Slide refSlide, bool coverShape = true)
        {
            if (refSlide == null)
            {
                return null;
            }

            // here we cut-paste the shape to get a reference of those shapes
            var oriShapeRange = refSlide.Shapes.Paste();

            // preprocess the shapes, eliminate animations for shapes
            foreach (Shape shape in oriShapeRange)
            {
                FromSlideFactory(refSlide).RemoveAnimationsForShape(shape);
            }

            // TODO: make use of PowerPointLabs.Presentation Model!!!
            // cut the original shape cover again and duplicate the slide
            // here the slide will be duplicated without the original shape cover
            oriShapeRange.Cut();
            var newSlide = FromSlideFactory(refSlide.Duplicate()[1]);
            
            // get a copy of original cover shapes
            var copyShapeRange = newSlide.Shapes.Paste();
            // paste the original shape cover back
            oriShapeRange = refSlide.Shapes.Paste();
            
            // make the range invisible before animated the slide
            copyShapeRange.Visible = Core.MsoTriState.msoFalse;

            MakeAnimatedBackground(newSlide);

            copyShapeRange.Visible = Core.MsoTriState.msoCTrue;
            oriShapeRange.Visible = Core.MsoTriState.msoCTrue;

            newSlide.Transition.EntryEffect = PpEntryEffect.ppEffectFadeSmoothly;
            newSlide.Transition.Duration = 0.5f;

            var bgEffectSlide = new PowerPointBgEffectSlide(newSlide.GetNativeSlide());

            try
            {
                if (coverShape)
                {
                    bgEffectSlide = PrepareForeground(oriShapeRange, copyShapeRange, refSlide, newSlide);
                }
            }
            catch (InvalidOperationException e)
            {
                refSlide.Delete();
                throw new InvalidOperationException(e.Message);
            }

            return bgEffectSlide;
        }
        # endregion

        # region API
        public void BlurAllBackground(int percentage)
        {
            AddBackgroundImage(null, percentage: percentage);
        }

        public void BlurBackground(int percentage, bool isTint)
        {
            AddBackgroundImage(null, percentage: percentage, isTint: isTint);
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
        private void AddBackgroundImage(IMatrixFilter filter, int percentage = 0, bool isTint = false)
        {
            if (filter == null)
            {
                EffectsLab.EffectsLabBlurSelected.BlurImage(AnimatedBackgroundPath, percentage);
            }
            else
            {
                using (var imageFactory = new ImageFactory())
                {
                    var image = imageFactory.Load(AnimatedBackgroundPath);

                    image = image.Filter(filter);

                    image.Save(AnimatedBackgroundPath);
                }
            }

            var newBackground = Shapes.AddPicture(AnimatedBackgroundPath, Core.MsoTriState.msoFalse,
                                                  Core.MsoTriState.msoTrue,
                                                  0, 0,
                                                  PowerPointPresentation.Current.SlideWidth,
                                                  PowerPointPresentation.Current.SlideHeight);

            newBackground.ZOrder(Core.MsoZOrderCmd.msoSendToBack);

            if (filter == null && isTint)
            {
                var overlayShape = EffectsLab.EffectsLabBlurSelected.GenerateOverlayShape(this, newBackground);
            }
        }

        private static Shape MakeFrontImage(ShapeRange shapeRange)
        {
            foreach (Shape shape in shapeRange)
            {
                shape.SoftEdge.Radius = Math.Min(Math.Min(shape.Width, shape.Height) * 0.15f, 10f);
            }

            var croppedShape = CropToShape.Crop(shapeRange, handleError: false);

            return croppedShape;
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

            var placeHolders =
                curSlide.Shapes.Cast<Shape>().Where(x => x.Type == Core.MsoShapeType.msoPlaceholder).ToList();

            foreach (var placeHolder in placeHolders)
            {
                placeHolder.Delete();
            }
        }

        private static PowerPointBgEffectSlide PrepareForeground(ShapeRange oriShapeRange, ShapeRange copyShapeRange,
                                                                 Slide refSlide, PowerPointSlide newSlide)
        {
            try
            {
                // crop in the original slide and put into clipboard
                var croppedShape = MakeFrontImage(oriShapeRange);

                croppedShape.Cut();

                // swap the uncropped shapes and cropped shapes
                var pastedCrop = newSlide.Shapes.Paste();

                // calibrate pasted shapes
                pastedCrop.Left -= 12;
                pastedCrop.Top -= 12;

                // ungroup front image if necessary
                if (pastedCrop[1].Type == Core.MsoShapeType.msoGroup)
                {
                    pastedCrop[1].Ungroup();
                }

                copyShapeRange.Cut();
                oriShapeRange = refSlide.Shapes.Paste();

                oriShapeRange.Fill.ForeColor.RGB = 0xaaaaaa;
                oriShapeRange.Fill.Transparency = 0.7f;
                oriShapeRange.Line.Visible = Core.MsoTriState.msoTrue;
                oriShapeRange.Line.ForeColor.RGB = 0x000000;

                Utils.Graphics.MakeShapeViewTimeInvisible(oriShapeRange, refSlide);

                oriShapeRange.Select();

                // finally add transition to the new slide
                newSlide.Transition.EntryEffect = PpEntryEffect.ppEffectFadeSmoothly;
                newSlide.Transition.Duration = 0.5f;

                return new PowerPointBgEffectSlide(newSlide.GetNativeSlide());
            }
            catch (Exception e)
            {
                var errorMessage = CropToShape.GetErrorMessageForErrorCode(e.Message);
                errorMessage = errorMessage.Replace("Crop To Shape", "Blur/Recolor Remainder");

                newSlide.Delete();

                throw new InvalidOperationException(errorMessage);
            }
        }
        # endregion
    }
}
