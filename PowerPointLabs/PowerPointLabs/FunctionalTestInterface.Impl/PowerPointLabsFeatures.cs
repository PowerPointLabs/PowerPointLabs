using System;
using FunctionalTestInterface;
using PowerPointLabs.FunctionalTestInterface.Impl.Controller;

namespace PowerPointLabs.FunctionalTestInterface.Impl
{
    [Serializable]
    class PowerPointLabsFeatures : MarshalByRefObject, IPowerPointLabsFeatures
    {
        private Ribbon1 Ribbon {
            get { return Globals.ThisAddIn.Ribbon; }
        }

        public void AutoCrop()
        {
            UIThreadExecutor.Execute(() =>
            {
                Ribbon.CropShapeButtonClick(new RibbonControl("AutoCrop"));
            });
        }

        public void AutoAnimate()
        {
            UIThreadExecutor.Execute(() =>
            {
                Ribbon.AddAnimationButtonClick(new RibbonControl("AutoAnimate"));
            });
        }

        public void AnimateInSlide()
        {
            UIThreadExecutor.Execute(() =>
            {
                Ribbon.AddInSlideAnimationButtonClick(new RibbonControl("AnimateInSlide"));
            });
        }

        public void AutoCaptions()
        {
            UIThreadExecutor.Execute(() =>
            {
                Ribbon.AddCaptionClick(new RibbonControl("AutoCaptions"));
            });
        }

        public void Spotlight()
        {
            UIThreadExecutor.Execute(() =>
            {
                Ribbon.SpotlightBtnClick(new RibbonControl("Spotlight"));
            });
        }

        public void FitToWidth()
        {
            UIThreadExecutor.Execute(() =>
            {
                Ribbon.FitToWidthClick(new RibbonControl("FitToWidth"));
            });
        }

        public void FitToHeight()
        {
            UIThreadExecutor.Execute(() =>
            {
                Ribbon.FitToHeightClick(new RibbonControl("FitToHeight"));
            });
        }

        public void ConvertToPic()
        {
            UIThreadExecutor.Execute(() =>
            {
                Ribbon.ConvertToPictureButtonClick(new RibbonControl("ConvertToPic"));
            });
        }

        public void DrillDown()
        {
            UIThreadExecutor.Execute(() =>
            {
                Ribbon.AddZoomInButtonClick(new RibbonControl("DrillDown"));
            });
        }

        public void StepBack()
        {
            UIThreadExecutor.Execute(() =>
            {
                Ribbon.AddZoomOutButtonClick(new RibbonControl("StepBack"));
            });
        }

        public void AddZoomToArea()
        {
            UIThreadExecutor.Execute(() =>
            {
                Ribbon.ZoomBtnClick(new RibbonControl("ZoomToArea"));
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
            UIThreadExecutor.Execute(() => {
                Ribbon.AddAudioClick(new RibbonControl("AutoNarrate"));
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

        public IColorsLabController ColorsLab
        {
            get { return ColorsLabController.Instance; }
        }

        public IShapesLabController ShapesLab
        {
            get { return ShapesLabController.Instance; }
        }
    }
}
