using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointLabs.Models
{
    class PowerPointZoomToAreaSingleSlide : PowerPointSlide
    {
        private PowerPoint.Shape indicatorShape = null;
        //private PowerPoint.Shape zoomSlidePicture = null;
        private PowerPoint.Shape zoomSlideCroppedShapes = null;
        //private float scaleFactorX = 0.0f;
        //private float scaleFactorY = 0.0f;

        private PowerPointZoomToAreaSingleSlide(PowerPoint.Slide slide) : base(slide)
        {
            _slide.Name = "PPTLabsMagnifyingSlide" + DateTime.Now.ToString("yyyyMMddHHmmssffff");
        }

        new public static PowerPointSlide FromSlideFactory(PowerPoint.Slide slide)
        {
            if (slide == null)
                return null;

            return new PowerPointZoomToAreaSingleSlide(slide);
        }

        public void PrepareForZoomToArea(List<PowerPoint.Shape> shapesToZoom)
        {
            MoveMotionAnimation();

            List<PowerPoint.Shape> shapes = _slide.Shapes.Cast<PowerPoint.Shape>().ToList();
            var matchingShapes = shapes.Where(current => (HasExitAnimation(current) || shapesToZoom.Contains(current)));
            foreach (PowerPoint.Shape s in matchingShapes)
                s.Delete();

            //_slide.Copy();
            //zoomSlidePicture = _slide.Shapes.PasteSpecial(PowerPoint.PpPasteDataType.ppPastePNG)[1];
            //zoomSlidePicture.Name = "PPTLabsZoomSlide" + DateTime.Now.ToString("yyyyMMddHHmmssffff");
            //scaleFactorX = PowerPointPresentation.SlideWidth / zoomSlidePicture.Width;
            //scaleFactorY = PowerPointPresentation.SlideHeight / zoomSlidePicture.Height;
            //FitShapeToSlide(ref zoomSlidePicture);

            AddZoomSlideCroppedPicture();

            DeleteSlideNotes();
            DeleteSlideMedia();
            ManageSlideTransitions();
            indicatorShape = AddPowerPointLabsIndicator();

            shapes = _slide.Shapes.Cast<PowerPoint.Shape>().ToList();
            matchingShapes = shapes.Where(current => (!(current.Equals(indicatorShape) || current.Equals(zoomSlideCroppedShapes))));
            foreach (PowerPoint.Shape s in matchingShapes)
            {
                DeleteShapeAnimations(s);
                PowerPoint.Effect effectFade = _slide.TimeLine.MainSequence.AddEffect(s, PowerPoint.MsoAnimEffect.msoAnimEffectFade, PowerPoint.MsoAnimateByLevel.msoAnimateLevelNone, PowerPoint.MsoAnimTriggerType.msoAnimTriggerWithPrevious);
                effectFade.Exit = Office.MsoTriState.msoTrue;
                effectFade.Timing.Duration = 0.25f;
            }
        }

        public void AddZoomToAreaAnimation(PowerPointSlide currentSlide, List<PowerPoint.Shape> shapesToZoom)
        {
            int shapeCount = 1;
            PowerPoint.Shape lastMagnifiedShape = null;
            foreach (PowerPoint.Shape zoomShape in shapesToZoom)
            {
                if (!ZoomToArea.backgroundZoomChecked)
                    ZoomWithoutBackground(zoomShape, shapeCount, ref lastMagnifiedShape, shapesToZoom.Count);
                else
                {
                }
                shapeCount++;
                zoomShape.Delete();
                indicatorShape.ZOrder(Office.MsoZOrderCmd.msoBringToFront);
            }
            zoomSlideCroppedShapes.Delete();
        }

        private void ZoomWithoutBackground(PowerPoint.Shape zoomShape, int shapeCount, ref PowerPoint.Shape lastMagnifiedShape, int totalShapes)
        {
            if (shapeCount == 1)
            {
                PowerPoint.Shape shapeToZoom = GetShapeToZoom(zoomShape);
                shapeToZoom.Name = "PPTLabsMagnifyingAreaSlide" + DateTime.Now.ToString("yyyyMMddHHmmssffff");
                PowerPoint.Shape referenceShape = GetReferenceShape(shapeToZoom);

                DefaultMotionAnimation.AddStepBackMotionAnimation(this, shapeToZoom, referenceShape, 0.5f, PowerPoint.MsoAnimTriggerType.msoAnimTriggerAfterPrevious);
                lastMagnifiedShape = GetLastMagnifiedShape(referenceShape);
                lastMagnifiedShape.Name = "PPTLabsMagnifyAreaGroupShape" + shapeCount + "-" + DateTime.Now.ToString("yyyyMMddHHmmssffff");

                PowerPoint.Effect effectAppear = _slide.TimeLine.MainSequence.AddEffect(lastMagnifiedShape, PowerPoint.MsoAnimEffect.msoAnimEffectAppear, PowerPoint.MsoAnimateByLevel.msoAnimateLevelNone, PowerPoint.MsoAnimTriggerType.msoAnimTriggerAfterPrevious);
                effectAppear.Timing.Duration = 0;

                PowerPoint.Effect effectDisappear = _slide.TimeLine.MainSequence.AddEffect(shapeToZoom, PowerPoint.MsoAnimEffect.msoAnimEffectAppear, PowerPoint.MsoAnimateByLevel.msoAnimateLevelNone, PowerPoint.MsoAnimTriggerType.msoAnimTriggerWithPrevious);
                effectDisappear.Exit = Office.MsoTriState.msoTrue;
                effectDisappear.Timing.Duration = 0;
            }
            else
            {
                PowerPoint.Shape tempShape1 = GetShapeToZoom(zoomShape);
                PowerPoint.Shape tempShape2 = GetReferenceShape(tempShape1);
                PowerPoint.Shape referenceShape = GetLastMagnifiedShape(tempShape2);
                tempShape1.Delete();
                referenceShape.Name = "PPTLabsMagnifyAreaGroupShape" + shapeCount + "-" + DateTime.Now.ToString("yyyyMMddHHmmssffff");
                PowerPoint.Shape shapeToZoom = lastMagnifiedShape;
                FrameMotionAnimation.animationType = FrameMotionAnimation.FrameMotionAnimationType.kZoomToAreaPan;
                FrameMotionAnimation.AddZoomToAreaPanFrameMotionAnimation(this, shapeToZoom, referenceShape);

                lastMagnifiedShape = referenceShape;
                PowerPoint.Effect effectAppear = _slide.TimeLine.MainSequence.AddEffect(lastMagnifiedShape, PowerPoint.MsoAnimEffect.msoAnimEffectAppear, PowerPoint.MsoAnimateByLevel.msoAnimateLevelNone, PowerPoint.MsoAnimTriggerType.msoAnimTriggerWithPrevious);
                effectAppear.Timing.Duration = 0;
            }

            if (shapeCount == totalShapes)
            {
                PowerPoint.Shape tempShape1 = GetShapeToZoom(zoomShape);
                PowerPoint.Shape shapeToZoom = GetReferenceShape(tempShape1);
                tempShape1.Delete();
                shapeToZoom.Name = "PPTLabsDeMagnifyAreaSlide" + shapeCount + "-" + DateTime.Now.ToString("yyyyMMddHHmmssffff");
                PowerPoint.Shape referenceShape = zoomShape;

                PowerPoint.Effect effectAppear = _slide.TimeLine.MainSequence.AddEffect(shapeToZoom, PowerPoint.MsoAnimEffect.msoAnimEffectAppear, PowerPoint.MsoAnimateByLevel.msoAnimateLevelNone, PowerPoint.MsoAnimTriggerType.msoAnimTriggerOnPageClick);
                effectAppear.Timing.Duration = 0;

                PowerPoint.Effect effectDisappear = _slide.TimeLine.MainSequence.AddEffect(lastMagnifiedShape, PowerPoint.MsoAnimEffect.msoAnimEffectAppear, PowerPoint.MsoAnimateByLevel.msoAnimateLevelNone, PowerPoint.MsoAnimTriggerType.msoAnimTriggerWithPrevious);
                effectDisappear.Timing.Duration = 0;
                effectDisappear.Exit = Office.MsoTriState.msoTrue;

                DefaultMotionAnimation.AddStepBackMotionAnimation(this, shapeToZoom, referenceShape, 0.5f, PowerPoint.MsoAnimTriggerType.msoAnimTriggerWithPrevious);
                ManageEndAnimations();
            }
        }

        private void ManageEndAnimations()
        {
            bool isFirst = true;
            PowerPoint.Effect effectFade = null;
            foreach (PowerPoint.Shape tmp in _slide.Shapes)
            {
                if (!(tmp.Equals(indicatorShape)) && !(tmp.Name.Contains("PPTLabsMagnifyAreaGroup")) && !(tmp.Name.Contains("PPTLabsMagnifyPanAreaGroup")) && !(tmp.Name.Contains("PPTLabsDeMagnifyAreaSlide")) && !(tmp.Name.Contains("PPTLabsMagnifyingAreaSlide")))
                {
                    if (isFirst)
                        effectFade = _slide.TimeLine.MainSequence.AddEffect(tmp, PowerPoint.MsoAnimEffect.msoAnimEffectFade, PowerPoint.MsoAnimateByLevel.msoAnimateLevelNone, PowerPoint.MsoAnimTriggerType.msoAnimTriggerAfterPrevious);
                    else
                        effectFade = _slide.TimeLine.MainSequence.AddEffect(tmp, PowerPoint.MsoAnimEffect.msoAnimEffectFade, PowerPoint.MsoAnimateByLevel.msoAnimateLevelNone, PowerPoint.MsoAnimTriggerType.msoAnimTriggerWithPrevious);
                    effectFade.Timing.Duration = 0.25f;
                    isFirst = false;
                }
            }
            isFirst = true;
            foreach (PowerPoint.Shape tmp in _slide.Shapes)
            {
                if (tmp.Name.Contains("PPTLabsMagnifyAreaGroup") || tmp.Name.Contains("PPTLabsMagnifyingAreaSlide") || tmp.Name.Contains("PPTLabsDeMagnifyAreaSlide"))
                {
                    if (isFirst)
                        effectFade = _slide.TimeLine.MainSequence.AddEffect(tmp, PowerPoint.MsoAnimEffect.msoAnimEffectFade, PowerPoint.MsoAnimateByLevel.msoAnimateLevelNone, PowerPoint.MsoAnimTriggerType.msoAnimTriggerAfterPrevious);
                    else
                        effectFade = _slide.TimeLine.MainSequence.AddEffect(tmp, PowerPoint.MsoAnimEffect.msoAnimEffectFade, PowerPoint.MsoAnimateByLevel.msoAnimateLevelNone, PowerPoint.MsoAnimTriggerType.msoAnimTriggerWithPrevious);

                    effectFade.Exit = Office.MsoTriState.msoTrue;
                    effectFade.Timing.Duration = 0.25f;
                    isFirst = false;
                }       
            }
        }

        private PowerPoint.Shape GetLastMagnifiedShape(PowerPoint.Shape referenceShape)
        {
            DeleteShapeAnimations(referenceShape);
            referenceShape.PictureFormat.CropLeft = 0;
            referenceShape.PictureFormat.CropTop = 0;
            referenceShape.PictureFormat.CropRight = 0;
            referenceShape.PictureFormat.CropBottom = 0;

            return referenceShape;
        }

        private PowerPoint.Shape GetReferenceShape(PowerPoint.Shape shapeToZoom)
        {
            shapeToZoom.Copy();

            PowerPoint.Shape referenceShape = _slide.Shapes.Paste()[1];
            referenceShape.LockAspectRatio = Office.MsoTriState.msoTrue;
            if (referenceShape.Width > referenceShape.Height)
                referenceShape.Width = PowerPointPresentation.SlideWidth;
            else
                referenceShape.Height = PowerPointPresentation.SlideHeight;

            referenceShape.Left = (PowerPointPresentation.SlideWidth / 2) - (referenceShape.Width / 2);
            referenceShape.Top = (PowerPointPresentation.SlideHeight / 2) - (referenceShape.Height / 2);
            
            return referenceShape;
        }

        private PowerPoint.Shape GetShapeToZoom(PowerPoint.Shape zoomShape)
        {
            PowerPoint.Shape shapeToZoom = zoomSlideCroppedShapes.Duplicate()[1];
            DeleteShapeAnimations(shapeToZoom);
            CopyShapePosition(zoomSlideCroppedShapes, ref shapeToZoom);

            shapeToZoom.PictureFormat.CropLeft += zoomShape.Left;
            shapeToZoom.PictureFormat.CropTop += zoomShape.Top;
            shapeToZoom.PictureFormat.CropRight += (PowerPointPresentation.SlideWidth - (zoomShape.Left + zoomShape.Width));
            shapeToZoom.PictureFormat.CropBottom += (PowerPointPresentation.SlideHeight - (zoomShape.Top + zoomShape.Height));

            CopyShapePosition(zoomShape, ref shapeToZoom);
            return shapeToZoom;
        }

        private void AddZoomSlideCroppedPicture()
        {
            PowerPointSlide zoomSlideCopy = this.Duplicate();
            Globals.ThisAddIn.Application.ActiveWindow.View.GotoSlide(zoomSlideCopy.Index);

            PowerPoint.Shape cropShape = zoomSlideCopy.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeRectangle, 0, 0, PowerPointPresentation.SlideWidth - 0.01f, PowerPointPresentation.SlideHeight - 0.01f);
            cropShape.Select();
            PowerPoint.Selection sel = Globals.ThisAddIn.Application.ActiveWindow.Selection;
            PowerPoint.Shape croppedShape = Globals.ThisAddIn.ribbon.CropShapeToSlide(ref sel);
            croppedShape.Cut();

            zoomSlideCroppedShapes = _slide.Shapes.PasteSpecial(PowerPoint.PpPasteDataType.ppPastePNG)[1];
            zoomSlideCroppedShapes.Name = "PPTLabsZoomGroup" + DateTime.Now.ToString("yyyyMMddHHmmssffff");
            //scaleFactorX = PowerPointPresentation.SlideWidth / zoomSlideCroppedShapes.Width;
            //scaleFactorY = PowerPointPresentation.SlideHeight / zoomSlideCroppedShapes.Height;
            FitShapeToSlide(ref zoomSlideCroppedShapes);
            zoomSlideCopy.Delete();
        }

        private void ManageSlideTransitions()
        {
            base.RemoveSlideTransitions();
            _slide.SlideShowTransition.AdvanceOnTime = Office.MsoTriState.msoFalse;
            _slide.SlideShowTransition.AdvanceOnClick = Office.MsoTriState.msoTrue;
        }

        private void FitShapeToSlide(ref PowerPoint.Shape shapeToMove)
        {
            shapeToMove.LockAspectRatio = Office.MsoTriState.msoFalse;
            shapeToMove.Left = 0;
            shapeToMove.Top = 0;
            shapeToMove.Width = PowerPointPresentation.SlideWidth;
            shapeToMove.Height = PowerPointPresentation.SlideHeight;
        }

        private static void CopyShapePosition(PowerPoint.Shape shapeToCopy, ref PowerPoint.Shape shapeToMove)
        {
            shapeToMove.Left = shapeToCopy.Left + (shapeToCopy.Width / 2) - (shapeToMove.Width / 2);
            shapeToMove.Top = shapeToCopy.Top + (shapeToCopy.Height / 2) - (shapeToMove.Height / 2);
        }
    }
}
