using System;
using FunctionalTestInterface;
using PowerPointLabs.FunctionalTestInterface.Impl.Controller;
using PowerPointLabs.Models;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointLabs.FunctionalTestInterface.Impl
{
    [Serializable]
    class PowerPointLabsFeatures : MarshalByRefObject, IPowerPointLabsFeatures
    {
        public void AutoCrop()
        {
            var selection = PowerPointCurrentPresentationInfo.CurrentSelection;
            CropToShape.Crop(selection);
        }

        public void AutoAnimate()
        {
            PowerPointLabs.AutoAnimate.AddAutoAnimation();
        }

        public void AnimateInSlide()
        {
            PowerPointLabs.AnimateInSlide.isHighlightBullets = false;
            PowerPointLabs.AnimateInSlide.AddAnimationInSlide();
        }

        public void AutoCaptions()
        {
            NotesToCaptions.EmbedCaptionsOnSelectedSlides();
        }

        public void Spotlight()
        {
            PowerPointLabs.Spotlight.AddSpotlightEffect();
        }

        public void FitToWidth()
        {
            var selectedShape = PowerPointCurrentPresentationInfo.CurrentSelection.ShapeRange[1];
            FitToSlide.FitToWidth(selectedShape);
        }

        public void FitToHeight()
        {
            var selectedShape = PowerPointCurrentPresentationInfo.CurrentSelection.ShapeRange[1];
            FitToSlide.FitToHeight(selectedShape);
        }

        public void ConvertToPic()
        {
            var selection = PowerPointCurrentPresentationInfo.CurrentSelection;
            ConvertToPicture.Convert(selection);
        }

        public void DrillDown()
        {
            AutoZoom.AddDrillDownAnimation();
        }

        public void StepBack()
        {
            AutoZoom.AddStepBackAnimation();
        }

        public void AddZoomToArea()
        {
            ZoomToArea.AddZoomToArea();
        }

        public void HighlightPoints()
        {
            if (Globals.ThisAddIn.Application.ActiveWindow.Selection.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
                HighlightBulletsText.userSelection = HighlightBulletsText.HighlightTextSelection.kShapeSelected;
            else if (Globals.ThisAddIn.Application.ActiveWindow.Selection.Type == PowerPoint.PpSelectionType.ppSelectionText)
                HighlightBulletsText.userSelection = HighlightBulletsText.HighlightTextSelection.kTextSelected;
            else
                HighlightBulletsText.userSelection = HighlightBulletsText.HighlightTextSelection.kNoneSelected;

            HighlightBulletsText.AddHighlightBulletsText();
        }

        public void HighlightBackground()
        {
            if (Globals.ThisAddIn.Application.ActiveWindow.Selection.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
                HighlightBulletsBackground.userSelection = HighlightBulletsBackground.HighlightBackgroundSelection.kShapeSelected;
            else if (Globals.ThisAddIn.Application.ActiveWindow.Selection.Type == PowerPoint.PpSelectionType.ppSelectionText)
                HighlightBulletsBackground.userSelection = HighlightBulletsBackground.HighlightBackgroundSelection.kTextSelected;
            else
                HighlightBulletsBackground.userSelection = HighlightBulletsBackground.HighlightBackgroundSelection.kNoneSelected;

            HighlightBulletsBackground.AddHighlightBulletsBackground();
        }

        public void HighlightFragments()
        {
            if (Globals.ThisAddIn.Application.ActiveWindow.Selection.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
                HighlightTextFragments.userSelection = HighlightTextFragments.HighlightTextSelection.kShapeSelected;
            else if (Globals.ThisAddIn.Application.ActiveWindow.Selection.Type == PowerPoint.PpSelectionType.ppSelectionText)
                HighlightTextFragments.userSelection = HighlightTextFragments.HighlightTextSelection.kTextSelected;
            else
                HighlightTextFragments.userSelection = HighlightTextFragments.HighlightTextSelection.kNoneSelected;

            HighlightTextFragments.AddHighlightedTextFragments();
        }

        public void GenerateTextAgenda()
        {
            AgendaLab.AgendaLabMain.GenerateAgenda(AgendaLab.Type.Bullet);
            GC.Collect();
        }

        public void GenerateVisualAgenda()
        {
            AgendaLab.AgendaLabMain.GenerateAgenda(AgendaLab.Type.Visual);
            GC.Collect();
        }

        public void GenerateBeamAgenda()
        {
            AgendaLab.AgendaLabMain.GenerateAgenda(AgendaLab.Type.Beam);
            GC.Collect();
        }

        public void RemoveAgenda()
        {
            AgendaLab.AgendaLabMain.RemoveAgenda();
            GC.Collect();
        }

        public void SynchronizeAgenda()
        {
            AgendaLab.AgendaLabMain.SynchroniseAgenda();
            GC.Collect();
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
