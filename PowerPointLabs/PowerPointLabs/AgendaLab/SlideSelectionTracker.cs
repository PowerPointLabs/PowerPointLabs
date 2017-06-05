using System;
using System.Collections.Generic;
using PowerPointLabs.Models;

namespace PowerPointLabs.AgendaLab
{
    /// <summary>
    /// Maintains the user's currently selected slides.
    /// Delete slides through the SlideSelectionTracker instead of just deleting them,
    /// so that SelectedSlides and UserCurrentSlides only point to existing slides.
    /// </summary>
    internal class SlideSelectionTracker
    {
#pragma warning disable 0618
        public PowerPointSlide UserCurrentSlide { get; private set; }
        private readonly List<PowerPointSlide> _selectedSlides;

        public List<PowerPointSlide> SelectedSlides
        {
            get { return new List<PowerPointSlide>(_selectedSlides); }
        }


        public SlideSelectionTracker(List<PowerPointSlide> selectedSlides, PowerPointSlide userCurrentSlide)
        {
            UserCurrentSlide = userCurrentSlide;
            _selectedSlides = selectedSlides;
        }

        public SlideSelectionTracker()
        {
        }

        /// <summary>
        /// Creates a "null" SlideSelectionTracker, Which deletes slides as usual, but without tracking.
        /// </summary>
        public static SlideSelectionTracker CreateInactiveTracker()
        {
            return new SlideSelectionTracker();
        }

        public void DeleteSlideAndTrack(PowerPointSlide slide)
        {
            RemoveSlideWithName(slide.Name);
            slide.Delete();
        }

        public void DeleteAcknowledgementSlideAndTrack()
        {
            RemoveSlideMeetingCondition(PowerPointAckSlide.IsAckSlide);
            PowerPointPresentation.Current.RemoveAckSlide();
        }

        private void RemoveSlideWithName(string name)
        {
            RemoveSlideMeetingCondition(slide => slide.Name == name);
        }

        private void RemoveSlideMeetingCondition(Predicate<PowerPointSlide> condition)
        {
            if (UserCurrentSlide != null && condition(UserCurrentSlide))
            {
                UserCurrentSlide = null;
            }

            _selectedSlides.RemoveAll(condition);
        }

    }
}
