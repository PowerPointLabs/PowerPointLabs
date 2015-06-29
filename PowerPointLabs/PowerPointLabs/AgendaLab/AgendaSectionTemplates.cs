using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using PowerPointLabs.Models;

namespace PowerPointLabs.AgendaLab
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
    /// 2) A pointer to a sync function, which is applied to synchronise the slide's contents.
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

        /// <exception cref="InvalidOperationException">Template already configured</exception>
        protected void AddConfiguration(AgendaSlideConfig[] frontSlides, AgendaSlideConfig[] backSlides)
        {
            if (_configured) throw new InvalidOperationException("Template already configured");

            FrontSlidesCount = frontSlides.Length;
            BackSlidesCount = backSlides.Length;
            FrontSlides = new ReadOnlyCollection<AgendaSlideConfig>(frontSlides);
            BackSlides = new ReadOnlyCollection<AgendaSlideConfig>(backSlides);

            _configured = true;
        }

        public abstract void ConfigHead();
        public abstract void ConfigMiddle();
        public abstract void ConfigEnd();
    }


    internal class VisualAgendaTemplate : AgendaTemplate
    {
        public override Type Type
        {
            get { return Type.Visual; }
        }

        public override void ConfigHead()
        {
            AgendaSlideConfig[] frontSlides = { };
            AgendaSlideConfig[] backSlides = { };
            AddConfiguration(frontSlides, backSlides);
        }

        public override void ConfigMiddle()
        {
            AgendaSlideConfig[] frontSlides = 
            {
                AgendaSlideConfig.AddSlide(AgendaLabMain.SyncVisualAgendaSlide, SlidePurpose.VisualAgendaSection),
            };

            AgendaSlideConfig[] backSlides = { };

            AddConfiguration(frontSlides, backSlides);
        }

        public override void ConfigEnd()
        {
            AgendaSlideConfig[] frontSlides = 
            {
                AgendaSlideConfig.AddSlide(AgendaLabMain.SyncVisualAgendaSlide, SlidePurpose.VisualAgendaSection),
            };

            AgendaSlideConfig[] backSlides = 
            {
                AgendaSlideConfig.AddSlide(AgendaLabMain.SyncVisualAgendaEndSlide, SlidePurpose.EndOfVisualAgenda),
            };

            AddConfiguration(frontSlides, backSlides);
        }
    }

    internal class BulletAgendaTemplate : AgendaTemplate
    {
        public override Type Type
        {
            get { return Type.Bullet; }
        }

        public override void ConfigHead()
        {
            AgendaSlideConfig[] frontSlides = { };

            AgendaSlideConfig[] backSlides = 
            {
                AgendaSlideConfig.AddSlide(AgendaLabMain.SyncBulletAgendaSlide, SlidePurpose.End),
            };

            AddConfiguration(frontSlides, backSlides);
        }

        public override void ConfigMiddle()
        {
            AgendaSlideConfig[] frontSlides = 
            {
                AgendaSlideConfig.AddSlide(AgendaLabMain.SyncBulletAgendaSlide, SlidePurpose.Start),
            };

            AgendaSlideConfig[] backSlides = 
            {
                AgendaSlideConfig.AddSlide(AgendaLabMain.SyncBulletAgendaSlide, SlidePurpose.End),
            };

            AddConfiguration(frontSlides, backSlides);
        }

        public override void ConfigEnd()
        {
            AgendaSlideConfig[] frontSlides = 
            {
                AgendaSlideConfig.AddSlide(AgendaLabMain.SyncBulletAgendaSlide, SlidePurpose.Start),
            };

            AgendaSlideConfig[] backSlides = 
            {
                AgendaSlideConfig.AddSlide(AgendaLabMain.SyncBulletAgendaSlide, SlidePurpose.End),
            };

            AddConfiguration(frontSlides, backSlides);
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

    /// <summary>
    /// Note: FrontIndexes and BackIndexes are intialised to NoSlide (-1).
    /// </summary>
    class TemplateIndexTable
    {
        public const int NoSlide = -1;
        public const int Reserved = -2;

        public readonly int[] FrontIndexes;
        public readonly int[] BackIndexes;

        public ReadOnlyCollection<PowerPointSlide> FrontSlideObjects;
        public ReadOnlyCollection<PowerPointSlide> BackSlideObjects;

        public TemplateIndexTable (int frontSlideCount, int backSlideCount)
        {
            FrontIndexes = new int[frontSlideCount];
            BackIndexes = new int[backSlideCount];
            for (int i = 0; i < FrontIndexes.Length; ++i) FrontIndexes[i] = NoSlide;
            for (int i = 0; i < BackIndexes.Length; ++i) BackIndexes[i] = NoSlide;
        }

        /// <summary>
        /// Stores the slide objects of the slides indexed by FrontIndexes and BackIndexes.
        /// </summary>
        public void StoreSlideObjects(List<PowerPointSlide> sectionSlides)
        {
            var frontSlideObjects = new PowerPointSlide[FrontIndexes.Length];
            var backSlideObjects = new PowerPointSlide[BackIndexes.Length];

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

    internal delegate void SyncFunction(PowerPointSlide refSlide,
                                        List<AgendaSection> sections,
                                        AgendaSection currentSection,
                                        List<string> deletedShapeNames,
                                        PowerPointSlide targetSlide);
}
