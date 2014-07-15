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
        private PowerPoint.Shape zoomSlideCroppedShapes = null;

        private PowerPointZoomToAreaSingleSlide(PowerPoint.Slide slide) : base(slide)
        {
            _slide.Name = "PPTLabsMagnifyingSingleSlide" + DateTime.Now.ToString("yyyyMMddHHmmssffff");
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

            //Delete zoom shapes and shapes with exit animations
            List<PowerPoint.Shape> shapes = _slide.Shapes.Cast<PowerPoint.Shape>().ToList();
            var matchingShapes = shapes.Where(current => (HasExitAnimation(current) || shapesToZoom.Contains(current)));
            foreach (PowerPoint.Shape s in matchingShapes)
                s.Delete();

            AddZoomSlideCroppedPicture();

            DeleteSlideNotes();
            DeleteSlideMedia();
            ManageSlideTransitions();
            indicatorShape = AddPowerPointLabsIndicator();

            //Add fade out effect for non-zoom shapes
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
                    ZoomWithBackground(zoomShape, shapeCount, ref lastMagnifiedShape, shapesToZoom.Count);
                shapeCount++;
                zoomShape.Delete();
                indicatorShape.ZOrder(Office.MsoZOrderCmd.msoBringToFront);
            }
            zoomSlideCroppedShapes.Delete();
        }
        private void ZoomWithBackground(PowerPoint.Shape zoomShape, int shapeCount, ref PowerPoint.Shape lastMagnifiedShape, int totalShapes)
        {
            if (shapeCount == 1)
            {
                //First Shape, add zoom in animation
                PowerPoint.Shape shapeToZoom = zoomSlideCroppedShapes.Duplicate()[1];
                PowerPointLabsGlobals.FitShapeToSlide(ref shapeToZoom);
                shapeToZoom.Name = "PPTLabsMagnifyingAreaSlide" + DateTime.Now.ToString("yyyyMMddHHmmssffff");

                PowerPoint.Shape referenceShape = GetReferenceShape(zoomShape);
                DefaultMotionAnimation.AddZoomToAreaMotionAnimation(this, shapeToZoom, zoomShape, referenceShape, 0.5f, PowerPoint.MsoAnimTriggerType.msoAnimTriggerAfterPrevious);

                referenceShape.Delete();
                PowerPoint.Shape tempShape1 = GetShapeToZoom(zoomShape);
                PowerPoint.Shape tempShape2 = GetReferenceShape(tempShape1);
                lastMagnifiedShape = GetLastMagnifiedShape(tempShape2);
                lastMagnifiedShape.Name = "PPTLabsMagnifyAreaGroupShape" + shapeCount + "-" + DateTime.Now.ToString("yyyyMMddHHmmssffff");
                tempShape1.Delete();

                PowerPoint.Effect effectAppear = _slide.TimeLine.MainSequence.AddEffect(lastMagnifiedShape, PowerPoint.MsoAnimEffect.msoAnimEffectAppear, PowerPoint.MsoAnimateByLevel.msoAnimateLevelNone, PowerPoint.MsoAnimTriggerType.msoAnimTriggerAfterPrevious);
                effectAppear.Timing.Duration = 0;

                PowerPoint.Effect effectDisappear = _slide.TimeLine.MainSequence.AddEffect(shapeToZoom, PowerPoint.MsoAnimEffect.msoAnimEffectAppear, PowerPoint.MsoAnimateByLevel.msoAnimateLevelNone, PowerPoint.MsoAnimTriggerType.msoAnimTriggerWithPrevious);
                effectDisappear.Exit = Office.MsoTriState.msoTrue;
                effectDisappear.Timing.Duration = 0;
            }
            else
            {
                //Add pan animation from last magnified shape
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
                //Last shape, add zoom out animation
                PowerPoint.Shape shapeToZoom = GetShapeToZoomWithBackground(zoomShape);

                PowerPoint.Effect effectAppear = _slide.TimeLine.MainSequence.AddEffect(shapeToZoom, PowerPoint.MsoAnimEffect.msoAnimEffectAppear, PowerPoint.MsoAnimateByLevel.msoAnimateLevelNone, PowerPoint.MsoAnimTriggerType.msoAnimTriggerOnPageClick);
                effectAppear.Timing.Duration = 0;

                FrameMotionAnimation.animationType = FrameMotionAnimation.FrameMotionAnimationType.kZoomToAreaDeMagnify;
                FrameMotionAnimation.AddStepBackFrameMotionAnimation(this, shapeToZoom);
                PowerPoint.Effect lastEffect = _slide.TimeLine.MainSequence[_slide.TimeLine.MainSequence.Count];
                ManageEndAnimationsForZoomWithBackground();
                lastEffect.MoveTo(_slide.TimeLine.MainSequence.Count);
                lastEffect.Timing.TriggerType = PowerPoint.MsoAnimTriggerType.msoAnimTriggerWithPrevious;
                lastEffect.Timing.TriggerDelayTime = 0.0f;
                lastEffect.Timing.Duration = 0.01f;
            }
        }

        private void ZoomWithoutBackground(PowerPoint.Shape zoomShape, int shapeCount, ref PowerPoint.Shape lastMagnifiedShape, int totalShapes)
        {
            if (shapeCount == 1)
            {
                //First Shape, add zoom in animation
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
                //Add pan animation from last magnified shape
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
                //Last shape, add zoom out animation
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
                ManageEndAnimationsForZoomWithoutBackground();
            }
        }

        //Add disappear animations for PPTLabs shapes and fade animation for remaining shapes
        private void ManageEndAnimationsForZoomWithBackground()
        {
            bool isFirst = true;
            PowerPoint.Effect effectAppear = null;
            foreach (PowerPoint.Shape tmp in _slide.Shapes)
            {
                if (!(tmp.Equals(indicatorShape)) && !(tmp.Name.Contains("PPTLabsMagnifyAreaGroup")) && !(tmp.Name.Contains("PPTLabsMagnifyPanAreaGroup")) && !(tmp.Name.Contains("PPTLabsDeMagnifyAreaSlide")) && !(tmp.Name.Contains("PPTLabsMagnifyingAreaSlide")))
                {
                    if (isFirst)
                        effectAppear = _slide.TimeLine.MainSequence.AddEffect(tmp, PowerPoint.MsoAnimEffect.msoAnimEffectAppear, PowerPoint.MsoAnimateByLevel.msoAnimateLevelNone, PowerPoint.MsoAnimTriggerType.msoAnimTriggerAfterPrevious);
                    else
                        effectAppear = _slide.TimeLine.MainSequence.AddEffect(tmp, PowerPoint.MsoAnimEffect.msoAnimEffectAppear, PowerPoint.MsoAnimateByLevel.msoAnimateLevelNone, PowerPoint.MsoAnimTriggerType.msoAnimTriggerWithPrevious);
                    effectAppear.Timing.Duration = 0.01f;
                    isFirst = false;
                }
                else if (tmp.Name.Contains("PPTLabsMagnifyAreaGroup") || tmp.Name.Contains("PPTLabsMagnifyingAreaSlide"))
                {
                    if (isFirst)
                        effectAppear = _slide.TimeLine.MainSequence.AddEffect(tmp, PowerPoint.MsoAnimEffect.msoAnimEffectAppear, PowerPoint.MsoAnimateByLevel.msoAnimateLevelNone, PowerPoint.MsoAnimTriggerType.msoAnimTriggerAfterPrevious);
                    else
                        effectAppear = _slide.TimeLine.MainSequence.AddEffect(tmp, PowerPoint.MsoAnimEffect.msoAnimEffectAppear, PowerPoint.MsoAnimateByLevel.msoAnimateLevelNone, PowerPoint.MsoAnimTriggerType.msoAnimTriggerWithPrevious);

                    effectAppear.Exit = Office.MsoTriState.msoTrue;
                    effectAppear.Timing.Duration = 0.01f;
                    isFirst = false;
                }
            }
        }

        //Add disappear animations for PPTLabs shapes and fade animation for remaining shapes
        private void ManageEndAnimationsForZoomWithoutBackground()
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
                        effectFade = _slide.TimeLine.MainSequence.AddEffect(tmp, PowerPoint.MsoAnimEffect.msoAnimEffectAppear, PowerPoint.MsoAnimateByLevel.msoAnimateLevelNone, PowerPoint.MsoAnimTriggerType.msoAnimTriggerAfterPrevious);
                    else
                        effectFade = _slide.TimeLine.MainSequence.AddEffect(tmp, PowerPoint.MsoAnimEffect.msoAnimEffectAppear, PowerPoint.MsoAnimateByLevel.msoAnimateLevelNone, PowerPoint.MsoAnimTriggerType.msoAnimTriggerWithPrevious);

                    effectFade.Exit = Office.MsoTriState.msoTrue;
                    effectFade.Timing.Duration = 0.01f;
                    isFirst = false;
                }       
            }
        }

        //Remove crop attributes of shape
        private PowerPoint.Shape GetLastMagnifiedShape(PowerPoint.Shape referenceShape)
        {
            DeleteShapeAnimations(referenceShape);
            referenceShape.PictureFormat.CropLeft = 0;
            referenceShape.PictureFormat.CropTop = 0;
            referenceShape.PictureFormat.CropRight = 0;
            referenceShape.PictureFormat.CropBottom = 0;

            return referenceShape;
        }

        //Return copy of argument shape that fits the slide
        private PowerPoint.Shape GetReferenceShape(PowerPoint.Shape shapeToZoom)
        {
            shapeToZoom.Copy();

            PowerPoint.Shape referenceShape = _slide.Shapes.Paste()[1];
            referenceShape.LockAspectRatio = Office.MsoTriState.msoTrue;
            if (referenceShape.Width > referenceShape.Height)
                referenceShape.Width = PowerPointCurrentPresentationInfo.SlideWidth;
            else
                referenceShape.Height = PowerPointCurrentPresentationInfo.SlideHeight;

            referenceShape.Left = (PowerPointCurrentPresentationInfo.SlideWidth / 2) - (referenceShape.Width / 2);
            referenceShape.Top = (PowerPointCurrentPresentationInfo.SlideHeight / 2) - (referenceShape.Height / 2);
            
            return referenceShape;
        }

        //Return cropped part of slide picture to be used for zoom in animation
        private PowerPoint.Shape GetShapeToZoom(PowerPoint.Shape zoomShape)
        {
            PowerPoint.Shape shapeToZoom = zoomSlideCroppedShapes.Duplicate()[1];
            DeleteShapeAnimations(shapeToZoom);
            PowerPointLabsGlobals.CopyShapePosition(zoomSlideCroppedShapes, ref shapeToZoom);

            shapeToZoom.PictureFormat.CropLeft += zoomShape.Left;
            shapeToZoom.PictureFormat.CropTop += zoomShape.Top;
            shapeToZoom.PictureFormat.CropRight += (PowerPointCurrentPresentationInfo.SlideWidth - (zoomShape.Left + zoomShape.Width));
            shapeToZoom.PictureFormat.CropBottom += (PowerPointCurrentPresentationInfo.SlideHeight - (zoomShape.Top + zoomShape.Height));

            PowerPointLabsGlobals.CopyShapePosition(zoomShape, ref shapeToZoom);
            return shapeToZoom;
        }

        //Return zoomed version of cropped slide picture to be used for zoom out animation
        private PowerPoint.Shape GetShapeToZoomWithBackground(PowerPoint.Shape zoomShape)
        {
            PowerPoint.Shape shapeToZoom = zoomSlideCroppedShapes.Duplicate()[1];
            PowerPointLabsGlobals.FitShapeToSlide(ref shapeToZoom);
            shapeToZoom.Name = "PPTLabsDeMagnifyAreaSlide" + DateTime.Now.ToString("yyyyMMddHHmmssffff");

            PowerPoint.Shape referenceShape = GetReferenceShape(zoomShape);

            float finalWidthMagnify = referenceShape.Width;
            float initialWidthMagnify = zoomShape.Width;
            float finalHeightMagnify = referenceShape.Height;
            float initialHeightMagnify = zoomShape.Height;

            zoomShape.Copy();
            PowerPoint.Shape zoomShapeCopy = _slide.Shapes.Paste()[1];
            PowerPointLabsGlobals.CopyShapeAttributes(zoomShape, ref zoomShapeCopy);

            Globals.ThisAddIn.Application.ActiveWindow.View.GotoSlide(_slide.SlideIndex);
            shapeToZoom.Select();
            zoomShapeCopy.Select(Office.MsoTriState.msoFalse);
            PowerPoint.ShapeRange selection = Globals.ThisAddIn.Application.ActiveWindow.Selection.ShapeRange;
            PowerPoint.Shape groupShape = selection.Group();

            groupShape.Width *= (finalWidthMagnify / initialWidthMagnify);
            groupShape.Height *= (finalHeightMagnify / initialHeightMagnify);
            groupShape.Ungroup();
            shapeToZoom.Left += (referenceShape.Left - zoomShapeCopy.Left);
            shapeToZoom.Top += (referenceShape.Top - zoomShapeCopy.Top);
            zoomShapeCopy.Delete();
            referenceShape.Delete();

            return shapeToZoom;
        }

        //Stores a cropped image of the slide in the global variable. Uses the CropShapeToSlide() method in Ribbon1.
        private void AddZoomSlideCroppedPicture()
        {
            PowerPointSlide zoomSlideCopy = this.Duplicate();
            Globals.ThisAddIn.Application.ActiveWindow.View.GotoSlide(zoomSlideCopy.Index);

            PowerPoint.Shape cropShape = zoomSlideCopy.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeRectangle, 0, 0, PowerPointCurrentPresentationInfo.SlideWidth - 0.01f, PowerPointCurrentPresentationInfo.SlideHeight - 0.01f);
            cropShape.Select();
            PowerPoint.Selection sel = Globals.ThisAddIn.Application.ActiveWindow.Selection;
            PowerPoint.Shape croppedShape = CropToShape.Crop(sel);
            croppedShape.Cut();

            zoomSlideCroppedShapes = _slide.Shapes.PasteSpecial(PowerPoint.PpPasteDataType.ppPastePNG)[1];
            zoomSlideCroppedShapes.Name = "PPTLabsZoomGroup" + DateTime.Now.ToString("yyyyMMddHHmmssffff");
            PowerPointLabsGlobals.FitShapeToSlide(ref zoomSlideCroppedShapes);
            zoomSlideCopy.Delete();
        }

        private void ManageSlideTransitions()
        {
            base.RemoveSlideTransitions();
            _slide.SlideShowTransition.AdvanceOnTime = Office.MsoTriState.msoFalse;
            _slide.SlideShowTransition.AdvanceOnClick = Office.MsoTriState.msoTrue;
        }
    }
}
