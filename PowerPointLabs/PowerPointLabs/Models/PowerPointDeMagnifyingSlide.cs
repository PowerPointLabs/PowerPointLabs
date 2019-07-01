using System;
using System.Linq;
using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.AnimationLab;
using PowerPointLabs.ZoomLab;

using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointLabs.Models
{
    class PowerPointDeMagnifyingSlide : PowerPointSlide
    {
#pragma warning disable 0618
        private PowerPoint.Shape indicatorShape = null;
        private PowerPoint.Shape zoomSlideCroppedShapes = null;
        private PowerPointDeMagnifyingSlide(PowerPoint.Slide slide) : base(slide)
        {
            _slide.Name = "PPTLabsDeMagnifyingSlide" + DateTime.Now.ToString("yyyyMMddHHmmssffff");
        }

        public static PowerPointSlide FromSlideFactory(PowerPoint.Slide slide)
        {
            if (slide == null)
            {
                return null;
            }

            return new PowerPointDeMagnifyingSlide(slide);
        }

        public void AddZoomToAreaAnimation(PowerPoint.Shape zoomShape)
        {
            PrepareForZoomToArea(zoomShape);
            
            if (!ZoomLabSettings.BackgroundZoomChecked)
            {
                //Zoom stored shape to fit slide
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

                DefaultMotionAnimation.AddDefaultMotionAnimation(this, zoomSlideCroppedShapes, zoomShape, 0.5f, PowerPoint.MsoAnimTriggerType.msoAnimTriggerWithPrevious);

                //Add appear animations to existing shapes
                bool isFirst = true;
                PowerPoint.Effect effectFade = null;
                foreach (PowerPoint.Shape tmp in _slide.Shapes)
                {
                    if (!(tmp.Equals(zoomSlideCroppedShapes) || tmp.Equals(indicatorShape)))
                    {
                        if (isFirst)
                        {
                            effectFade = _slide.TimeLine.MainSequence.AddEffect(tmp, PowerPoint.MsoAnimEffect.msoAnimEffectFade, PowerPoint.MsoAnimateByLevel.msoAnimateLevelNone, PowerPoint.MsoAnimTriggerType.msoAnimTriggerAfterPrevious);
                        }
                        else
                        {
                            effectFade = _slide.TimeLine.MainSequence.AddEffect(tmp, PowerPoint.MsoAnimEffect.msoAnimEffectFade, PowerPoint.MsoAnimateByLevel.msoAnimateLevelNone, PowerPoint.MsoAnimTriggerType.msoAnimTriggerWithPrevious);
                        }

                        effectFade.Timing.Duration = 0.25f;
                        isFirst = false;
                    }
                }

                //Add fade out anmation to shape added by PPTLabs
                effectFade = _slide.TimeLine.MainSequence.AddEffect(zoomSlideCroppedShapes, PowerPoint.MsoAnimEffect.msoAnimEffectFade, PowerPoint.MsoAnimateByLevel.msoAnimateLevelNone, PowerPoint.MsoAnimTriggerType.msoAnimTriggerWithPrevious);
                effectFade.Exit = Office.MsoTriState.msoTrue;
                effectFade.Timing.Duration = 0.25f;
            }
            else
            {
                GetShapeToZoomWithBackground(zoomShape);
                PowerPoint.Effect lastDisappearEffect = DefaultMotionAnimation.AddZoomOutMotionAnimation(this,
                    zoomSlideCroppedShapes, PowerPoint.MsoAnimTriggerType.msoAnimTriggerWithPrevious);
                DefaultMotionAnimation.PreloadShape(this, zoomSlideCroppedShapes);
                
                //Add appear animations to existing shapes
                bool isFirst = true;
                PowerPoint.Effect effectFade = null;
                foreach (PowerPoint.Shape tmp in _slide.Shapes)
                {
                    if (!(tmp.Equals(zoomSlideCroppedShapes) || tmp.Equals(indicatorShape)) && !(tmp.Name.Contains("PPTLabsMagnifyShape")) && !(tmp.Name.Contains("PPTLabsMagnifyArea")))
                    {
                        tmp.Visible = Office.MsoTriState.msoTrue;
                        if (isFirst)
                        {
                            effectFade = _slide.TimeLine.MainSequence.AddEffect(tmp, PowerPoint.MsoAnimEffect.msoAnimEffectAppear, PowerPoint.MsoAnimateByLevel.msoAnimateLevelNone, PowerPoint.MsoAnimTriggerType.msoAnimTriggerAfterPrevious);
                        }
                        else
                        {
                            effectFade = _slide.TimeLine.MainSequence.AddEffect(tmp, PowerPoint.MsoAnimEffect.msoAnimEffectAppear, PowerPoint.MsoAnimateByLevel.msoAnimateLevelNone, PowerPoint.MsoAnimTriggerType.msoAnimTriggerWithPrevious);
                        }

                        effectFade.Timing.Duration = 0.01f;
                        isFirst = false;
                    }
                }
                
                //Move last frame disappear animation to end 
                lastDisappearEffect.MoveTo(_slide.TimeLine.MainSequence.Count);
                lastDisappearEffect.Timing.TriggerType = PowerPoint.MsoAnimTriggerType.msoAnimTriggerWithPrevious;
                lastDisappearEffect.Timing.TriggerDelayTime = 0.01f;
            }

            indicatorShape.ZOrder(Office.MsoZOrderCmd.msoBringToFront);
        }

        private void PrepareForZoomToArea(PowerPoint.Shape zoomShape)
        {
            RemoveAnimationsForShapes(_slide.Shapes.Cast<PowerPoint.Shape>().ToList());
            DeleteIndicator();
            DeleteShapesWithPrefix("PPTLabsMagnifyAreaSlide");

            AddZoomSlideCroppedPicture(zoomShape);

            DeleteSlideNotes();
            DeleteSlideMedia();
            ManageSlideTransitions();
            indicatorShape = AddPowerPointLabsIndicator();
        }

        //Return zoomed version of cropped slide picture to be used for zoom out animation
        private void GetShapeToZoomWithBackground(PowerPoint.Shape zoomShape)
        {
            PowerPoint.Shape referenceShape = GetReferenceShape(zoomShape);

            float finalWidthMagnify = referenceShape.Width;
            float initialWidthMagnify = zoomShape.Width;
            float finalHeightMagnify = referenceShape.Height;
            float initialHeightMagnify = zoomShape.Height;

            zoomShape.Copy();
            PowerPoint.Shape zoomShapeCopy = _slide.Shapes.Paste()[1];
            LegacyShapeUtil.CopyShapeAttributes(zoomShape, ref zoomShapeCopy);

            Globals.ThisAddIn.Application.ActiveWindow.View.GotoSlide(_slide.SlideIndex);
            zoomSlideCroppedShapes.Select();
            zoomShapeCopy.Visible = Office.MsoTriState.msoTrue;
            zoomShapeCopy.Select(Office.MsoTriState.msoFalse);
            PowerPoint.ShapeRange selection = Globals.ThisAddIn.Application.ActiveWindow.Selection.ShapeRange;
            PowerPoint.Shape groupShape = selection.Group();

            groupShape.Width *= (finalWidthMagnify / initialWidthMagnify);
            groupShape.Height *= (finalHeightMagnify / initialHeightMagnify);
            groupShape.Ungroup();
            zoomSlideCroppedShapes.Left += (referenceShape.Left - zoomShapeCopy.Left);
            zoomSlideCroppedShapes.Top += (referenceShape.Top - zoomShapeCopy.Top);
            zoomShapeCopy.SafeDelete();
            referenceShape.SafeDelete();
        }

        private PowerPoint.Shape GetReferenceShape(PowerPoint.Shape shapeToZoom)
        {
            PowerPoint.Shape referenceShape = shapeToZoom.Duplicate()[1];
            referenceShape.LockAspectRatio = Office.MsoTriState.msoTrue;
            if (referenceShape.Width > referenceShape.Height)
            {
                referenceShape.Width = PowerPointPresentation.Current.SlideWidth;
            }
            else
            {
                referenceShape.Height = PowerPointPresentation.Current.SlideHeight;
            }

            referenceShape.Left = (PowerPointPresentation.Current.SlideWidth / 2) - (referenceShape.Width / 2);
            referenceShape.Top = (PowerPointPresentation.Current.SlideHeight / 2) - (referenceShape.Height / 2);

            return referenceShape;
        }

        //Store cropped version of slide picture as global variable
        private void AddZoomSlideCroppedPicture(PowerPoint.Shape zoomShape)
        {
            zoomSlideCroppedShapes = GetShapesWithPrefix("PPTLabsMagnifyAreaGroup")[0];
            zoomSlideCroppedShapes.Visible = Office.MsoTriState.msoTrue;
            DeleteShapeAnimations(zoomSlideCroppedShapes);

            if (!ZoomLabSettings.BackgroundZoomChecked)
            {
                zoomSlideCroppedShapes.PictureFormat.CropLeft += zoomShape.Left;
                zoomSlideCroppedShapes.PictureFormat.CropTop += zoomShape.Top;
                zoomSlideCroppedShapes.PictureFormat.CropRight += (PowerPointPresentation.Current.SlideWidth - (zoomShape.Left + zoomShape.Width));
                zoomSlideCroppedShapes.PictureFormat.CropBottom += (PowerPointPresentation.Current.SlideHeight - (zoomShape.Top + zoomShape.Height));

                LegacyShapeUtil.CopyShapePosition(zoomShape, ref zoomSlideCroppedShapes);
            }
        }

        private void ManageSlideTransitions()
        {
            RemoveSlideTransitions();
            _slide.SlideShowTransition.AdvanceOnTime = Office.MsoTriState.msoFalse;
            _slide.SlideShowTransition.AdvanceOnClick = Office.MsoTriState.msoTrue;
        }
    }
}
