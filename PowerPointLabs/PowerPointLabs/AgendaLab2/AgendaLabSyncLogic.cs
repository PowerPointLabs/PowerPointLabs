using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using PowerPointLabs.Models;
using PowerPointLabs.Utils;
using PowerPointLabs.Views;
using Graphics = PowerPointLabs.Utils.Graphics;
using Shape = Microsoft.Office.Interop.PowerPoint.Shape;

namespace PowerPointLabs.AgendaLab2
{
    internal static partial class AgendaLabMain
    {
        private static void ConfigureTemplate(AgendaSection section, AgendaTemplate template)
        {
            if (section.Index == 1)
                template.ConfigHead();
            else if (section.Index == NumberOfSections)
                template.ConfigEnd();
            else
                template.ConfigMiddle();
        }

        private static void l(params Object[] lines)
        {
            foreach (var line in lines)
            {
                Debug.WriteLine(line.ToString());
            }
        }

        /// <summary>
        /// Rebuilds the slides in the section to match the slides indicated by the template.
        /// Does not rename the agendaslides.
        /// Assumption: Reference slide is the first slide.
        /// </summary>
        private static void RebuildSectionUsingTemplate(AgendaSection section, AgendaTemplate template)
        {
            ConfigureTemplate(section, template);

            // Step 1: Generate Assignment List and fill in Template Table.
            var templateTable = template.CreateIndexTable();
            var sectionSlides = GetSectionSlides(section);
            if (AgendaSlide.IsReferenceslide(sectionSlides[0])) sectionSlides.RemoveAt(0);

            var addToIndex = SectionLastSlideIndex(section) + 1;

            var assignmentList = new List<int>();
            for (var i = 0; i < sectionSlides.Count; ++i) assignmentList.Add(-1);

            // Step 1a: Filling in Template Table
            MatchTemplateTableWithSlides(template, sectionSlides, templateTable, assignmentList);

            // Step 1b: Generating Assignment List
            int indexOfFirstBackSlide;
            GenerateInitialAssignmentList(out indexOfFirstBackSlide, template, templateTable, assignmentList, sectionSlides);

            // Step 2: Add all missing slides.
            var createdSlides = AddAllMissingSlides(ref addToIndex, template, templateTable, assignmentList, indexOfFirstBackSlide);
            sectionSlides.AddRange(createdSlides);

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
            markedForDeletion.ForEach(slide => slide.Delete());
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
        private static List<PowerPointSlide> AddAllMissingSlides(ref int addToIndex, AgendaTemplate template, TemplateIndexTable templateTable, List<int> assignmentList, int indexOfFirstBackSlide)
        {
            var createdSlides = new List<PowerPointSlide>();

            for (int i = 0; i < template.FrontSlidesCount; ++i)
            {
                if (templateTable.FrontIndexes[i] == TemplateIndexTable.NoSlide)
                {
                    var newSlide = AddBlankSlide(addToIndex);
                    createdSlides.Add(newSlide);
                    AgendaSlide.SetSlideName(newSlide, Type.None, template.FrontSlides[i].SlidePurpose, AgendaSection.None);
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
                    AgendaSlide.SetSlideName(newSlide, Type.None, template.BackSlides[i].SlidePurpose, AgendaSection.None);
                    assignmentList.Add(indexOfFirstBackSlide + i);
                    addToIndex++;
                }
            }

            return createdSlides;
        }

        /// <summary>
        /// Assignment list specs after generation:
        /// TemplateIndexTable.NoSlide: Marked for deletion
        /// any other index: The new position of the slide. [relative to startIndex. first slide is index 0.]
        /// </summary>
        private static void GenerateInitialAssignmentList(out int indexOfFirstBackSlide, AgendaTemplate template,
            TemplateIndexTable templateTable, List<int> assignmentList, List<PowerPointSlide> sectionSlides)
        {
            for (int i = 0; i < template.FrontSlidesCount; ++i)
            {
                int chosenSlide = templateTable.FrontIndexes[i];
                if (chosenSlide == -1) continue;
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
                if (chosenSlide == -1) continue;
                assignmentList[chosenSlide] = indexOfFirstBackSlide + i;
            }
        }

        private static void MatchTemplateTableWithSlides(AgendaTemplate template, List<PowerPointSlide> sectionSlides,
            TemplateIndexTable templateTable, List<int> assignmentList)
        {
            for (int i = 0; i < template.FrontSlidesCount; ++i)
            {
                var purpose = template.FrontSlides[i].SlidePurpose;
                for (int j = 0; j < assignmentList.Count; ++j)
                {
                    if (!AgendaSlide.MatchesPurpose(sectionSlides[j], purpose)) continue;
                    l("MATCH " + purpose);
                    templateTable.FrontIndexes[i] = j;
                    assignmentList[j] = TemplateIndexTable.Reserved;
                    break;
                }
            }

            for (int i = 0; i < template.BackSlidesCount; ++i)
            {
                var purpose = template.BackSlides[i].SlidePurpose;
                for (int j = 0; j < assignmentList.Count; ++j)
                {
                    if (!AgendaSlide.MatchesPurpose(sectionSlides[j], purpose)) continue;
                    l("MATCHBACK " + purpose);
                    templateTable.BackIndexes[i] = j;
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
    }
}
