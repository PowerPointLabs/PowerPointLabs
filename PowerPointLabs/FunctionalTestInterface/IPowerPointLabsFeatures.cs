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
        void SetZoomProperties(bool backgroundChecked, bool multiSlideChecked);
        void HighlightPoints();
        void HighlightBackground();
        void HighlightFragments();
        void AutoNarrate();

        // Agenda Lab
        void GenerateTextAgenda();
        void GenerateVisualAgenda();
        void GenerateBeamAgenda();
        void RemoveAgenda();
        void SynchronizeAgenda();

        // Effects Lab
        void TransparentEffect();
        void MagnifyingGlassEffect();
        void BlurRemainderEffect();
        void GreyScaleEffect();
        void BlackAndWhiteEffect();
        void GothamEffect();
        void SepiaEffect();

        IColorsLabController ColorsLab { get; }
        IShapesLabController ShapesLab { get; }
    }
}
