using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.Office.Interop.PowerPoint;
using PowerPointLabs.AgendaLab.Templates;
using PowerPointLabs.Models;
using PowerPointLabs.Utils;

namespace PowerPointLabs.AgendaLab
{
    internal static partial class AgendaLabMain
    {
#pragma warning disable 0618
        #region Main Synchronisation Function
        /// <summary>
        /// Call the function like this for example:
        /// SynchroniseSlidesUsingTemplate(slideTracker, refSlide, () => new VisualAgendaTemplate());
        /// generateTemplate is a function that returns a newly created template.
        /// </summary>
        private static void SynchroniseSlidesUsingTemplate(SlideSelectionTracker slideTracker, PowerPointSlide refSlide, Func<AgendaTemplate> generateTemplate)
        {
            var sections = Sections;

            var deletedShapeNames = RetrieveTrackedDeletions(refSlide);

            refSlide.DeleteSlideNumberShapes();
            refSlide.MakeShapeNamesNonDefault();
            refSlide.MakeShapeNamesUnique(shape => !AgendaShape.IsAnyAgendaShape(shape) &&
                                                   !PowerPointSlide.IsTemplateSlideMarker(shape));

            ScrambleSlideSectionNames();
            foreach (var currentSection in sections)
            {
                var template = generateTemplate();
                ConfigureTemplate(currentSection, template);

                var templateTable = RebuildSectionUsingTemplate(slideTracker, currentSection, template);
                SynchroniseAllSlides(template, templateTable, refSlide, sections, deletedShapeNames, currentSection);
            }

            TrackShapesInSlide(refSlide);
        }

        private static void ConfigureTemplate(AgendaSection section, AgendaTemplate template)
        {
            if (section.Index == 1)
            {
                template.ConfigHead();
            }
            else if (section.Index == NumberOfSections)
            {
                template.ConfigEnd();
            }
            else
            {
                template.ConfigMiddle();
            }
        }
        #endregion


        #region Rebuilding the arrangement of slides in a Section
        /// <summary>
        /// Rebuilds the slides in the section to match the slides indicated by the template.
        /// Names the agenda slides properly.
        /// Assumption: Reference slide is the first slide.
        /// </summary>
        private static TemplateIndexTable RebuildSectionUsingTemplate(SlideSelectionTracker slideTracker, AgendaSection currentSection, AgendaTemplate template)
        {
            if (template.NotConfigured)
            {
                throw new ArgumentException("Template is not configured yet.");
            }

            // Step 1: Generate Assignment List and fill in Template Table.
            var templateTable = template.CreateIndexTable();
            var sectionSlides = GetSectionSlides(currentSection);
            if (AgendaSlide.IsReferenceslide(sectionSlides[0]))
            {
                sectionSlides.RemoveAt(0);
            }

            var addToIndex = SectionLastSlideIndex(currentSection) + 1;

            var assignmentList = new List<int>();
            for (var i = 0; i < sectionSlides.Count; ++i)
            {
                assignmentList.Add(-1);
            }

            // Step 1a: Filling in Template Table
            MatchTemplateTableWithSlides(template, sectionSlides, templateTable, assignmentList, currentSection);

            // Step 1b: Generating Assignment List
            int indexOfFirstBackSlide;
            GenerateInitialAssignmentList(out indexOfFirstBackSlide, template, templateTable, assignmentList, sectionSlides);

            // Step 2: Add all missing slides.
            var createdSlides = AddAllMissingSlides(ref addToIndex, template, templateTable, assignmentList, currentSection, indexOfFirstBackSlide);
            sectionSlides.AddRange(createdSlides);
            templateTable.StoreSlideObjects(sectionSlides);


            // Step 3: Create Goal Array of Slide Objects and MarkedForDeletion list.
            List<PowerPointSlide> markedForDeletion;
            int newSlideCount = indexOfFirstBackSlide + template.BackSlidesCount;
            var goalArray = GenerateGoalArray(newSlideCount, assignmentList, sectionSlides, out markedForDeletion);

            // Step 4: Use goal array to reorder all goal slides.
            foreach (var slide in goalArray)
            {
                slide.MoveTo(addToIndex-1);
            }

            // Step 5: Delete all slides marked for deletion.
            markedForDeletion.ForEach(slideTracker.DeleteSlideAndTrack);


            return templateTable;
        }

        private static PowerPointSlide[] GenerateGoalArray(int newSlideCount, List<int> assignmentList,
            List<PowerPointSlide> sectionSlides, out List<PowerPointSlide> markedForDeletion)
        {
            PowerPointSlide[] goalArray = new PowerPointSlide[newSlideCount];
            markedForDeletion = new List<PowerPointSlide>();
            for (int i = 0; i < assignmentList.Count; ++i)
            {
                int assignedIndex = assignmentList[i];
                if (assignedIndex == -1)
                {
                    markedForDeletion.Add(sectionSlides[i]);
                }
                else
                {
                    goalArray[assignedIndex] = sectionSlides[i];
                }
            }
            return goalArray;
        }

        /// <summary>
        /// Returns a list of the newly added slides.
        /// Updates assignmentList (by appending)
        /// Gives placeholder agendaslide name to the created slides.
        /// </summary>
        private static List<PowerPointSlide> AddAllMissingSlides(ref int addToIndex, AgendaTemplate template,
            TemplateIndexTable templateTable, List<int> assignmentList, AgendaSection currentSection,
            int indexOfFirstBackSlide)
        {
            var createdSlides = new List<PowerPointSlide>();

            for (int i = 0; i < template.FrontSlidesCount; ++i)
            {
                if (templateTable.FrontIndexes[i] == TemplateIndexTable.NoSlide)
                {
                    var newSlide = AddBlankSlide(addToIndex);
                    createdSlides.Add(newSlide);
                    AgendaSlide.SetSlideName(newSlide, template.Type, template.FrontSlides[i].SlidePurpose,
                        currentSection);

                    templateTable.IsNewlyGeneratedFront[i] = true;
                    templateTable.FrontIndexes[i] = assignmentList.Count;
                    assignmentList.Add(i);
                    addToIndex++;
                }
            }
            for (int i = 0; i < template.BackSlidesCount; ++i)
            {
                if (templateTable.BackIndexes[i] == TemplateIndexTable.NoSlide)
                {
                    var newSlide = AddBlankSlide(addToIndex);
                    createdSlides.Add(newSlide);
                    AgendaSlide.SetSlideName(newSlide, template.Type, template.BackSlides[i].SlidePurpose,
                        currentSection);

                    templateTable.IsNewlyGeneratedBack[i] = true;
                    templateTable.BackIndexes[i] = assignmentList.Count;
                    assignmentList.Add(indexOfFirstBackSlide + i);
                    addToIndex++;
                }
            }

            return createdSlides;
        }

        /// <summary>
        /// The assignment list indicates the new position of each of the previous slides.
        /// assignmentList[oldSlideIndex] = newSlideIndex
        /// 
        /// if newSlideIndex is equal to TemplateIndexTable.NoSlide, it means the slide is marked for deletion.
        /// 
        /// All slideIndexes are relative to the index of the first slide in the section. first slide is index 0.
        /// </summary>
        private static void GenerateInitialAssignmentList(out int indexOfFirstBackSlide, AgendaTemplate template,
            TemplateIndexTable templateTable, List<int> assignmentList, List<PowerPointSlide> sectionSlides)
        {
            for (int i = 0; i < template.FrontSlidesCount; ++i)
            {
                int chosenSlide = templateTable.FrontIndexes[i];
                if (chosenSlide == -1)
                {
                    continue;
                }
                assignmentList[chosenSlide] = i;
            }
            int currentIndex = template.FrontSlidesCount;
            for (int i = 0; i < assignmentList.Count; ++i)
            {
                if (assignmentList[i] == TemplateIndexTable.NoSlide)
                {
                    if (!AgendaSlide.IsAnyAgendaSlide(sectionSlides[i]))
                    {
                        assignmentList[i] = currentIndex;
                        currentIndex++;
                    }
                }
            }
            indexOfFirstBackSlide = currentIndex;
            for (int i = 0; i < template.BackSlidesCount; ++i)
            {
                int chosenSlide = templateTable.BackIndexes[i];
                if (chosenSlide == -1)
                {
                    continue;
                }
                assignmentList[chosenSlide] = indexOfFirstBackSlide + i;
            }
        }

        private static void MatchTemplateTableWithSlides(AgendaTemplate template, List<PowerPointSlide> sectionSlides,
            TemplateIndexTable templateTable, List<int> assignmentList, AgendaSection currentSection)
        {
            for (int i = 0; i < template.FrontSlidesCount; ++i)
            {
                var purpose = template.FrontSlides[i].SlidePurpose;
                for (int j = 0; j < assignmentList.Count; ++j)
                {
                    if (!AgendaSlide.MatchesPurpose(sectionSlides[j], purpose))
                    {
                        continue;
                    }
                    templateTable.FrontIndexes[i] = j;
                    AgendaSlide.SetSlideName(sectionSlides[j], template.Type, purpose, currentSection);
                    assignmentList[j] = TemplateIndexTable.Reserved;
                    break;
                }
            }

            for (int i = 0; i < template.BackSlidesCount; ++i)
            {
                var purpose = template.BackSlides[i].SlidePurpose;
                for (int j = 0; j < assignmentList.Count; ++j)
                {
                    if (!AgendaSlide.MatchesPurpose(sectionSlides[j], purpose))
                    {
                        continue;
                    }
                    templateTable.BackIndexes[i] = j;
                    AgendaSlide.SetSlideName(sectionSlides[j], template.Type, purpose, currentSection);
                    assignmentList[j] = TemplateIndexTable.Reserved;
                    break;
                }
            }
        }

        private static PowerPointSlide AddBlankSlide(int addLocation)
        {
            var slide =
                PowerPointSlide.FromSlideFactory(PowerPointPresentation.Current
                                                                       .Presentation
                                                                       .Slides
                                                                       .Add(addLocation, PpSlideLayout.ppLayoutBlank),
                                                                        includeIndicator: true);
            return slide;
        }

        #endregion


        #region Synchronising the content of slides

        private static void SynchroniseAllSlides(AgendaTemplate template, TemplateIndexTable templateTable,
            PowerPointSlide refSlide, List<AgendaSection> sections, List<string> deletedShapeNames, AgendaSection currentSection)
        {
            if (template.NotConfigured)
            {
                throw new ArgumentException("Template is not configured yet.");
            }

            for (int i = 0; i < template.FrontSlidesCount; ++i)
            {
                template.FrontSlides[i].SyncFunction(refSlide, sections, currentSection, deletedShapeNames,
                    templateTable.IsNewlyGeneratedFront[i], templateTable.FrontSlideObjects[i]);
            }
            for (int i = 0; i < template.BackSlidesCount; ++i)
            {
                template.BackSlides[i].SyncFunction(refSlide, sections, currentSection, deletedShapeNames,
                    templateTable.IsNewlyGeneratedBack[i], templateTable.BackSlideObjects[i]);
            }
        }

        /// <summary>
        /// Retrieves tracked shapes from the slide metadata, finds out which shapes have been deleted by the user,
        /// and returns the names of those deleted shapes.
        /// </summary>
        private static List<string> RetrieveTrackedDeletions(PowerPointSlide slide)
        {
            var retrievedNameList = Common.UnserializeCollection(slide.RetrieveDataFromNotes());
            if (retrievedNameList == null)
            {
                return new List<string>();
            }

            var currentNames = slide.GetNameToShapeDictionary();
            return retrievedNameList.Where(name => !currentNames.ContainsKey(name)).ToList();
        }

        /// <summary>
        /// Tracks shapes in the slide metadata. (used for tracking deletions later on)
        /// </summary>
        private static void TrackShapesInSlide(PowerPointSlide slide)
        {
            var nameList = slide.GetNameToShapeDictionary().Keys.ToList();
            slide.StoreDataInNotes(Common.SerializeCollection(nameList));
        }
        #endregion
    }
}
