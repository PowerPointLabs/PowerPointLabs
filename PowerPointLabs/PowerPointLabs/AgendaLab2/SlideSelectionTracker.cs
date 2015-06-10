using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using PowerPointLabs.Models;

namespace PowerPointLabs.AgendaLab2
{
    /// <summary>
    /// Maintains the user's currently selected slides.
    /// Delete slides through the SlideSelectionTracker instead of just deleting them,
    /// so that SelectedSlides and UserCurrentSlides only point to existing slides.
    /// </summary>
    internal class SlideSelectionTracker
    {
        private readonly bool isActive;

        public PowerPointSlide UserCurrentSlide { get; private set; }
        private List<PowerPointSlide> _SelectedSlides;

        public List<PowerPointSlide> SelectedSlides
        {
            get { return new List<PowerPointSlide>(SelectedSlides); }
        }


        public SlideSelectionTracker(List<PowerPointSlide> selectedSlides, PowerPointSlide userCurrentSlide)
        {
            isActive = true;
            UserCurrentSlide = userCurrentSlide;
            _SelectedSlides = selectedSlides;
        }

        public SlideSelectionTracker()
        {
            isActive = false;
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

            if (condition(UserCurrentSlide)) UserCurrentSlide = null;
            _SelectedSlides.RemoveAll(condition);
        }

    }
}
