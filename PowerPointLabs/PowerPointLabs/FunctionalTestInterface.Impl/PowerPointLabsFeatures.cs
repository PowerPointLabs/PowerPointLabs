using System;
using System.Drawing;

using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.FunctionalTestInterface.Impl.Controller;
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

        public IHighlightLabController HighlightLab
        {
            get { return HighlightLabController.Instance; }
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
                var control = new RibbonControl("CropToShape");
                Ribbon.OnAction(control);
            });
        }

        public void CropOutPadding()
        {
            UIThreadExecutor.Execute(() =>
            {
                var control = new RibbonControl("CropOutPaddingButton");
                Ribbon.OnAction(control);
            });
        }

        public void CropToAspectRatioW1H10()
        {
            UIThreadExecutor.Execute(() =>
            {
                var control = new RibbonControl("CropToAspectRatioOption1_10");
                control.Tag = "CropToAspectRatio";
                Ribbon.OnAction(control);
            });
        }

        public void CropToSlide()
        {
            var control = new RibbonControl("CropToSlideButton");
            Ribbon.OnAction(control);
        }

        public void CropToSame()
        {
            var control = new RibbonControl("CropToSameButton");
            Ribbon.OnAction(control);
        }

        public void AutoAnimate()
        {
            UIThreadExecutor.Execute(() =>
            {
                var control = new RibbonControl("AddAnimationSlide");
                Ribbon.OnAction(control);
            });
        }

        public void AnimateInSlide()
        {
            UIThreadExecutor.Execute(() =>
            {
                var control = new RibbonControl("AnimateInSlide");
                Ribbon.OnAction(control);
            });
        }

        public void AutoCaptions()
        {
            UIThreadExecutor.Execute(() =>
            {
                Ribbon.OnAction(new RibbonControl("AddCaptions"));
            });
        }

        public void Spotlight()
        {
            UIThreadExecutor.Execute(() =>
            {
                Ribbon.SpotlightBtnClick(new RibbonControl("Spotlight"));
            });
        }

        public void SetSpotlightProperties(float newTransparency, float newSoftEdge, Color newColor)
        {
            UIThreadExecutor.Execute(() =>
            {
                Ribbon.SpotlightPropertiesEdited(newTransparency, newSoftEdge, newColor);
            });
        }

        public void OpenSpotlightDialog()
        {
            UIThreadExecutor.Execute(() =>
            {
                Ribbon.SpotlightDialogButtonPressed(new RibbonControl("OpenSpotlightDialog"));
            });
        }

        public void ConvertToPic()
        {
            UIThreadExecutor.Execute(() =>
            {
                Ribbon.OnAction(new RibbonControl("ConvertToPicture"));
            });
        }

        public void DrillDown()
        {
            UIThreadExecutor.Execute(() =>
            {
                Ribbon.OnAction(new RibbonControl("DrillDown"));
            });
        }

        public void StepBack()
        {
            UIThreadExecutor.Execute(() =>
            {
                Ribbon.OnAction(new RibbonControl("StepBack"));
            });
        }

        public void AddZoomToArea()
        {
            UIThreadExecutor.Execute(() =>
            {
                Ribbon.OnAction(new RibbonControl("ZoomToArea"));
            });
        }

        public void SetZoomProperties(bool backgroundChecked, bool multiSlideChecked)
        {
            UIThreadExecutor.Execute(() =>
            {
                Ribbon.ZoomPropertiesEdited(backgroundChecked, multiSlideChecked);
            });
        }

        public void HighlightPoints()
        {
            UIThreadExecutor.Execute(() =>
            {
                Ribbon.HighlightBulletsTextButtonClick(new RibbonControl("HighlightPoints"));
            });
        }

        public void HighlightBackground()
        {
            UIThreadExecutor.Execute(() =>
            {
                Ribbon.HighlightBulletsBackgroundButtonClick(new RibbonControl("HighlightBackground"));
            });
        }

        public void HighlightFragments()
        {
            UIThreadExecutor.Execute(() =>
            {
                Ribbon.HighlightTextFragmentsButtonClick(new RibbonControl("HighlightFragments"));
            });
        }

        public void AutoNarrate()
        {
            UIThreadExecutor.Execute(() =>
            {
                Ribbon.OnAction(new RibbonControl("AddNarrations"));
            });
        }

        public void GenerateTextAgenda()
        {
            UIThreadExecutor.Execute(() =>
            {
                Ribbon.BulletPointAgendaClick(new RibbonControl("TextAgenda"));
            });
        }

        public void GenerateVisualAgenda()
        {
            UIThreadExecutor.Execute(() =>
            {
                Ribbon.VisualAgendaClick(new RibbonControl("VisualAgenda"));
            });
        }

        public void GenerateBeamAgenda()
        {
            UIThreadExecutor.Execute(() =>
            {
                Ribbon.BeamAgendaClick(new RibbonControl("BeamAgenda"));
            });
        }

        public void RemoveAgenda()
        {
            UIThreadExecutor.Execute(() =>
            {
                Ribbon.RemoveAgendaClick(new RibbonControl("RemoveAgenda"));
            });
        }

        public void SynchronizeAgenda()
        {
            UIThreadExecutor.Execute(() =>
            {
                Ribbon.UpdateAgendaClick(new RibbonControl("SyncAgenda"));
            });
        }

        public void TransparentEffect()
        {
            UIThreadExecutor.Execute(() =>
            {
                Ribbon.TransparentEffectClick(new RibbonControl("TransparentEffect"));
            });
        }

        public void MagnifyingGlassEffect()
        {
            UIThreadExecutor.Execute(() =>
            {
                Ribbon.MagnifyGlassEffectClick(new RibbonControl("MagnifyingGlassEffect"));
            });
        }
        
        public void BlurrinessOverlay(string feature, bool pressed)
        {
            UIThreadExecutor.Execute(() =>
            {
                var control = new RibbonControl(feature + TextCollection.DynamicMenuCheckBoxId);
                control.Tag = "Blurriness";
                Ribbon.OnCheckBoxAction(control, pressed);
            });
        }

        public void BlurSelectedEffect()
        {
            UIThreadExecutor.Execute(() =>
            {
                var control = new RibbonControl("EffectsLabBlurSelectedOption90");
                control.Tag = "Blurriness";
                Ribbon.OnAction(control);
            });
        }

        public void BlurRemainderEffect()
        {
            UIThreadExecutor.Execute(() =>
            {
                var control = new RibbonControl("EffectsLabBlurRemainderOption90");
                control.Tag = "Blurriness";
                Ribbon.OnAction(control);
            });
        }

        public void GreyScaleRemainderEffect()
        {
            UIThreadExecutor.Execute(() =>
            {
                Ribbon.GreyScaleRemainderEffectClick(new RibbonControl("GreyScaleEffect"));
            });
        }


        public void GothamRemainderEffect()
        {
            UIThreadExecutor.Execute(() =>
            {
                Ribbon.GothamRemainderEffectClick(new RibbonControl("GothamEffect"));
            });
        }

        public void SepiaRemainderEffect()
        {
            UIThreadExecutor.Execute(() =>
            {
                Ribbon.SepiaRemainderEffectClick(new RibbonControl("SepiaEffect"));
            });
        }


        public void BlurBackgroundEffect()
        {
            UIThreadExecutor.Execute(() =>
            {
                var control = new RibbonControl("EffectsLabBlurBackgroundOption90");
                control.Tag = "Blurriness";
                Ribbon.OnAction(control);
            });
        }

        public void BlackAndWhiteBackgroundEffect()
        {
            UIThreadExecutor.Execute(() =>
            {
                Ribbon.BlackWhiteBackgroundEffectClick(new RibbonControl("BlackAndWhiteEffect"));
            });
        }

        public void SepiaBackgroundEffect()
        {
            UIThreadExecutor.Execute(() =>
            {
                Ribbon.SepiaBackgroundEffectClick(new RibbonControl("SepiaEffect"));
            });
        }

        public void PasteToFillSlide()
        {
            UIThreadExecutor.Execute(() =>
            {
                var control = new RibbonControl("PasteToFillSlide");
                Ribbon.OnAction(control);
            });
        }

        public void PasteAtCursorPosition()
        {
            UIThreadExecutor.Execute(() =>
            {
                var control = new RibbonControl("PasteAtCursorPosition");
                Ribbon.OnAction(control);
            });
        }

        public void PasteAtOriginalPosition()
        {
            UIThreadExecutor.Execute(() =>
            {
                var control = new RibbonControl("PasteAtOriginalPosition");
                Ribbon.OnAction(control);
            });
        }

        public void PasteIntoGroup()
        {
            UIThreadExecutor.Execute(() =>
            {
                var control = new RibbonControl("PasteIntoGroup");
                Ribbon.OnAction(control);
            });
        }

        public void ReplaceWithClipboard()
        {
            UIThreadExecutor.Execute(() =>
            {
                var control = new RibbonControl("ReplaceWithClipboard");
                Ribbon.OnAction(control);
            });
        }
    }
}
