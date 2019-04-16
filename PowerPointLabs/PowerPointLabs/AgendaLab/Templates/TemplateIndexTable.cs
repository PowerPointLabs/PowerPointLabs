using System.Collections.Generic;
using System.Collections.ObjectModel;

using PowerPointLabs.Models;

namespace PowerPointLabs.AgendaLab.Templates
{
    /// <summary>
    /// Note: FrontIndexes and BackIndexes are intialised to NoSlide (-1).
    /// </summary>
    class TemplateIndexTable
    {
        public const int NoSlide = -1;
        public const int Reserved = -2;

        public readonly int[] FrontIndexes;
        public readonly int[] BackIndexes;

        public readonly bool[] IsNewlyGeneratedFront;
        public readonly bool[] IsNewlyGeneratedBack;

        public ReadOnlyCollection<PowerPointSlide> FrontSlideObjects;
        public ReadOnlyCollection<PowerPointSlide> BackSlideObjects;

        public TemplateIndexTable(int frontSlideCount, int backSlideCount)
        {
            FrontIndexes = new int[frontSlideCount];
            BackIndexes = new int[backSlideCount];
            for (int i = 0; i < FrontIndexes.Length; ++i)
            {
                FrontIndexes[i] = NoSlide;
            }

            for (int i = 0; i < BackIndexes.Length; ++i)
            {
                BackIndexes[i] = NoSlide;
            }

            IsNewlyGeneratedFront = new bool[frontSlideCount];
            IsNewlyGeneratedBack = new bool[backSlideCount];
        }

        /// <summary>
        /// Stores the slide objects of the slides indexed by FrontIndexes and BackIndexes.
        /// </summary>
        public void StoreSlideObjects(List<PowerPointSlide> sectionSlides)
        {
            PowerPointSlide[] frontSlideObjects = new PowerPointSlide[FrontIndexes.Length];
            PowerPointSlide[] backSlideObjects = new PowerPointSlide[BackIndexes.Length];

            for (int i = 0; i < FrontIndexes.Length; ++i)
            {
                frontSlideObjects[i] = sectionSlides[FrontIndexes[i]];
            }
            for (int i = 0; i < BackIndexes.Length; ++i)
            {
                backSlideObjects[i] = sectionSlides[BackIndexes[i]];
            }

            FrontSlideObjects = new ReadOnlyCollection<PowerPointSlide>(frontSlideObjects);
            BackSlideObjects = new ReadOnlyCollection<PowerPointSlide>(backSlideObjects);
        }
    }
}
