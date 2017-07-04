using System;
using System.Drawing;

using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.EffectsLab;
using PowerPointLabs.FunctionalTestInterface.Impl.Controller;
using PowerPointLabs.ZoomLab;
using TestInterface;

namespace PowerPointLabs.FunctionalTestInterface.Impl
{
    [Serializable]
    class PowerPointLabsFeatures : MarshalByRefObject, IPowerPointLabsFeatures
    {
        public IColorsLabController ColorsLab
        {
            get { return ColorsLabController.Instance; }
        }

        public IShapesLabController ShapesLab
        {
            get { return ShapesLabController.Instance; }
        }

        public IPositionsLabController PositionsLab
        {
            get { return PositionsLabController.Instance; }
        }

        public ISyncLabController SyncLab
        {
            get { return SyncLabController.Instance; }
        }

        public ITimerLabController TimerLab
        {
            get { return TimerLabController.Instance; }
        }

        private Ribbon1 Ribbon
        {
            get { return FunctionalTestExtensions.GetRibbonUi(); }
        }

        public void AutoCrop()
        {
            UIThreadExecutor.Execute(() =>
            {
                Ribbon.OnAction(new RibbonControl(TextCollection.CropToShapeTag));
            });
        }

        public void CropOutPadding()
        {
            UIThreadExecutor.Execute(() =>
            {
                Ribbon.OnAction(new RibbonControl(TextCollection.CropOutPaddingTag));
            });
        }

        public void CropToAspectRatioW1H10()
        {
            UIThreadExecutor.Execute(() =>
            {
                var control = new RibbonControl("CropToAspectRatioOption1_10");
                control.Tag = TextCollection.CropToAspectRatioTag;
                Ribbon.OnAction(control);
            });
        }

        public void CropToSlide()
        {
            Ribbon.OnAction(new RibbonControl(TextCollection.CropToSlideTag));
        }

        public void CropToSame()
        {
            Ribbon.OnAction(new RibbonControl(TextCollection.CropToSameDimensionsTag));
        }

        public void AutoAnimate()
        {
            UIThreadExecutor.Execute(() =>
            {
                Ribbon.OnAction(new RibbonControl(TextCollection.AddAnimationSlideTag));
            });
        }

        public void AnimateInSlide()
        {
            UIThreadExecutor.Execute(() =>
            {
                Ribbon.OnAction(new RibbonControl(TextCollection.AnimateInSlideTag));
            });
        }

        public void AutoCaptions()
        {
            UIThreadExecutor.Execute(() =>
            {
                Ribbon.OnAction(new RibbonControl(TextCollection.AddCaptionsTag));
            });
        }

        public void Spotlight()
        {
            UIThreadExecutor.Execute(() =>
            {
                Ribbon.OnAction(new RibbonControl("AddSpotlight"));
            });
        }

        public void SetSpotlightProperties(float newTransparency, float newSoftEdge, Color newColor)
        {
            UIThreadExecutor.Execute(() =>
            {
                EffectsLabSpotlightSettings.SpotlightPropertiesEdited(newTransparency, newSoftEdge, newColor);
            });
        }

        public void OpenSpotlightDialog()
        {
            UIThreadExecutor.Execute(() =>
            {
                Ribbon.OnAction(new RibbonControl("SpotlightSettings"));
            });
        }

        public void ConvertToPic()
        {
            UIThreadExecutor.Execute(() =>
            {
                Ribbon.OnAction(new RibbonControl(TextCollection.ConvertToPictureTag));
            });
        }

        public void DrillDown()
        {
            UIThreadExecutor.Execute(() =>
            {
                Ribbon.OnAction(new RibbonControl(TextCollection.DrillDownTag));
            });
        }

        public void StepBack()
        {
            UIThreadExecutor.Execute(() =>
            {
                Ribbon.OnAction(new RibbonControl(TextCollection.StepBackTag));
            });
        }

        public void AddZoomToArea()
        {
            UIThreadExecutor.Execute(() =>
            {
                Ribbon.OnAction(new RibbonControl(TextCollection.ZoomToAreaTag));
            });
        }

        public void SetZoomProperties(bool backgroundChecked, bool multiSlideChecked)
        {
            UIThreadExecutor.Execute(() =>
            {
                ZoomLabSettings.ZoomLabSettingsEdited(backgroundChecked, multiSlideChecked);
            });
        }

        public void HighlightPoints()
        {
            UIThreadExecutor.Execute(() =>
            {
                Ribbon.OnAction(new RibbonControl(TextCollection.HighlightPointsTag));
            });
        }

        public void HighlightBackground()
        {
            UIThreadExecutor.Execute(() =>
            {
                Ribbon.OnAction(new RibbonControl(TextCollection.HighlightBackgroundTag));
            });
        }

        public void HighlightFragments()
        {
            UIThreadExecutor.Execute(() =>
            {
                Ribbon.OnAction(new RibbonControl(TextCollection.HighlightTextTag));
            });
        }

        public void RemoveHighlight()
        {
            UIThreadExecutor.Execute(() =>
            {
                Ribbon.OnAction(new RibbonControl(TextCollection.RemoveHighlightTag));
            });
        }

        public void AutoNarrate()
        {
            UIThreadExecutor.Execute(() =>
            {
                Ribbon.OnAction(new RibbonControl(TextCollection.AddNarrationsTag));
            });
        }

        public void GenerateTextAgenda()
        {
            UIThreadExecutor.Execute(() =>
            {
                Ribbon.OnAction(new RibbonControl(TextCollection.TextAgendaTag));
            });
        }

        public void GenerateVisualAgenda()
        {
            UIThreadExecutor.Execute(() =>
            {
                Ribbon.OnAction(new RibbonControl(TextCollection.VisualAgendaTag));
            });
        }

        public void GenerateBeamAgenda()
        {
            UIThreadExecutor.Execute(() =>
            {
                Ribbon.OnAction(new RibbonControl(TextCollection.BeamAgendaTag));
            });
        }

        public void RemoveAgenda()
        {
            UIThreadExecutor.Execute(() =>
            {
                Ribbon.OnAction(new RibbonControl(TextCollection.RemoveAgendaTag));
            });
        }

        public void SynchronizeAgenda()
        {
            UIThreadExecutor.Execute(() =>
            {
                Ribbon.OnAction(new RibbonControl(TextCollection.UpdateAgendaTag));
            });
        }

        public void TransparentEffect()
        {
            UIThreadExecutor.Execute(() =>
            {
                Ribbon.OnAction(new RibbonControl(TextCollection.MakeTransparentTag));
            });
        }

        public void MagnifyingGlassEffect()
        {
            UIThreadExecutor.Execute(() =>
            {
                Ribbon.OnAction(new RibbonControl(TextCollection.MagnifyTag));
            });
        }
        
        public void BlurrinessOverlay(string feature, bool pressed)
        {
            UIThreadExecutor.Execute(() =>
            {
                var control = new RibbonControl(feature + TextCollection.DynamicMenuCheckBoxId);
                control.Tag = TextCollection.EffectsLabBlurrinessTag;
                Ribbon.OnCheckBoxAction(control, pressed);
            });
        }

        public void BlurSelectedEffect()
        {
            UIThreadExecutor.Execute(() =>
            {
                var control = new RibbonControl("BlurSelectedOption90");
                control.Tag = TextCollection.EffectsLabBlurrinessTag;
                Ribbon.OnAction(control);
            });
        }

        public void BlurRemainderEffect()
        {
            UIThreadExecutor.Execute(() =>
            {
                var control = new RibbonControl("BlurRemainderOption90");
                control.Tag = TextCollection.EffectsLabBlurrinessTag;
                Ribbon.OnAction(control);
            });
        }

        public void GreyScaleRemainderEffect()
        {
            UIThreadExecutor.Execute(() =>
            {
                var control = new RibbonControl("GrayScaleRecolorRemainderMenu");
                control.Tag = TextCollection.EffectsLabRecolorTag;
                Ribbon.OnAction(control);
            });
        }


        public void GothamRemainderEffect()
        {
            UIThreadExecutor.Execute(() =>
            {
                var control = new RibbonControl("GothamRecolorRemainderMenu");
                control.Tag = TextCollection.EffectsLabRecolorTag;
                Ribbon.OnAction(control);
            });
        }

        public void SepiaRemainderEffect()
        {
            UIThreadExecutor.Execute(() =>
            {
                var control = new RibbonControl("SepiaRecolorRemainderMenu");
                control.Tag = TextCollection.EffectsLabRecolorTag;
                Ribbon.OnAction(control);
            });
        }


        public void BlurBackgroundEffect()
        {
            UIThreadExecutor.Execute(() =>
            {
                var control = new RibbonControl("BlurBackgroundOption90");
                control.Tag = TextCollection.EffectsLabBlurrinessTag;
                Ribbon.OnAction(control);
            });
        }

        public void BlackAndWhiteBackgroundEffect()
        {
            UIThreadExecutor.Execute(() =>
            {
                var control = new RibbonControl("BlackAndWhiteRecolorBackgroundMenu");
                control.Tag = TextCollection.EffectsLabRecolorTag;
                Ribbon.OnAction(control);
            });
        }

        public void SepiaBackgroundEffect()
        {
            UIThreadExecutor.Execute(() =>
            {
                var control = new RibbonControl("SepiaRecolorBackgroundMenu");
                control.Tag = TextCollection.EffectsLabRecolorTag;
                Ribbon.OnAction(control);
            });
        }

        public void PasteToFillSlide()
        {
            UIThreadExecutor.Execute(() =>
            {
                Ribbon.OnAction(new RibbonControl(TextCollection.PasteToFillSlideTag));
            });
        }

        public void PasteAtCursorPosition()
        {
            UIThreadExecutor.Execute(() =>
            {
                Ribbon.OnAction(new RibbonControl(TextCollection.PasteAtCursorPositionTag));
            });
        }

        public void PasteAtOriginalPosition()
        {
            UIThreadExecutor.Execute(() =>
            {
                Ribbon.OnAction(new RibbonControl(TextCollection.PasteAtOriginalPositionTag));
            });
        }

        public void PasteIntoGroup()
        {
            UIThreadExecutor.Execute(() =>
            {
                Ribbon.OnAction(new RibbonControl(TextCollection.PasteIntoGroupTag));
            });
        }

        public void ReplaceWithClipboard()
        {
            UIThreadExecutor.Execute(() =>
            {
                Ribbon.OnAction(new RibbonControl(TextCollection.ReplaceWithClipboardTag));
            });
        }
    }
}
