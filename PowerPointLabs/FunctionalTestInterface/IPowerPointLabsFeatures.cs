namespace FunctionalTestInterface
{
    public interface IPowerPointLabsFeatures
    {
        void AutoCrop();
        void AutoAnimate();
        void AnimateInSlide();
        void AutoCaptions();
        void Spotlight();
        void FitToWidth();
        void FitToHeight();
        void ConvertToPic();
        void DrillDown();
        void StepBack();
        void AddZoomToArea();
        void HighlightPoints();
        void HighlightBackground();
        void HighlightFragments();

        // Agenda Lab
        void GenerateTextAgenda();
        void GenerateVisualAgenda();
        void GenerateBeamAgenda();
        void RemoveAgenda();
        void SynchronizeAgenda();
    }
}
