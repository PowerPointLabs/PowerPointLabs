using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;

using PowerPointLabs.Models;

namespace PowerPointLabs.AgendaLab.Templates
{
    /// <summary>
    /// A template for defining the positions of generated slides for each section in the agenda.
    /// Does not include the reference slide.
    /// 
    /// There are three types of section for a template: Head, Middle and End.
    /// Head refers to the first section. End refers to the last section. Middle refers to all other sections.
    /// 
    /// Each configuration must have the following:
    /// 1) The SlidePurpose of the slide.
    /// 2) A pointer to a sync function, which is applied to synchronize the slide's contents.
    /// 
    /// Template Rules:
    /// - Every generated slide in the section must have a different purpose.
    /// </summary>
    internal abstract class AgendaTemplate
    {
        public abstract Type Type { get; }
        public int FrontSlidesCount { get; private set; }
        public int BackSlidesCount { get; private set; }
        public ReadOnlyCollection<AgendaSlideConfig> FrontSlides { get; private set; }
        public ReadOnlyCollection<AgendaSlideConfig> BackSlides { get; private set; }
        private bool _configured;

        public TemplateIndexTable CreateIndexTable()
        {
            return new TemplateIndexTable(FrontSlidesCount, BackSlidesCount);
        }

        public bool NotConfigured
        {
            get { return !_configured; }
        }

        public abstract void ConfigHead();
        public abstract void ConfigMiddle();
        public abstract void ConfigEnd();

        /// <exception cref="InvalidOperationException">Template already configured</exception>
        protected void AddConfiguration(AgendaSlideConfig[] frontSlides, AgendaSlideConfig[] backSlides)
        {
            if (_configured)
            {
                throw new InvalidOperationException("Template already configured");
            }

            FrontSlidesCount = frontSlides.Length;
            BackSlidesCount = backSlides.Length;
            FrontSlides = new ReadOnlyCollection<AgendaSlideConfig>(frontSlides);
            BackSlides = new ReadOnlyCollection<AgendaSlideConfig>(backSlides);

            _configured = true;
        }
    }

    struct AgendaSlideConfig
    {
        public readonly SyncFunction SyncFunction;
        public readonly SlidePurpose SlidePurpose;

        private AgendaSlideConfig(SyncFunction syncFunction, SlidePurpose slidePurpose)
        {
            SyncFunction = syncFunction;
            SlidePurpose = slidePurpose;
        }

        public static AgendaSlideConfig AddSlide(SyncFunction syncFunction, SlidePurpose slidePurpose)
        {
            return new AgendaSlideConfig(syncFunction, slidePurpose);
        }
    }

    internal delegate void SyncFunction(PowerPointSlide refSlide,
                                        List<AgendaSection> sections,
                                        AgendaSection currentSection,
                                        List<string> deletedShapeNames,
                                        bool isNewlyGenerated,
                                        PowerPointSlide targetSlide);
}
