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
    /// <summary>
    /// The sections should not change during generation / syncing.
    /// </summary>
    internal static partial class AgendaLabMain
    {
        private static LoadingDialog _loadDialog = new LoadingDialog();

        #region Bullet Formats
        private struct BulletFormats
        {
            public readonly TextRange2 Visited;
            public readonly TextRange2 Highlighted;
            public readonly TextRange2 Unvisited;

            public static BulletFormats ExtractFormats(Shape contentShape)
            {
                throw new NotImplementedException();
            }
        }


        #endregion

        #region API
        public static void GenerateAgenda(Type type)
        {
            bool dialogOpen = false;
            var curWindow = Globals.ThisAddIn.Application.ActiveWindow;
            var oldViewType = curWindow.ViewType;

            try
            {
                var slideTracker = new SlideSelectionTracker(SelectedSlides, CurrentSlide);

                if (AgendaPresent())
                {
                    var confirm = MessageBox.Show(TextCollection.AgendaLabAgendaExistError,
                                                  TextCollection.AgendaLabAgendaExistErrorCaption,
                                                  MessageBoxButtons.OKCancel);
                    if (confirm != DialogResult.OK) return;

                    RemoveAllAgendaItems(slideTracker);
                }

                if (!ValidSections()) return;

                slideTracker.DeleteAcknowledgementSlideAndTrack();

                dialogOpen = DisplayLoadingDialog(TextCollection.AgendaLabLoadingDialogTitle,
                                                    TextCollection.AgendaLabLoadingDialogContent);
                curWindow.ViewType = PpViewType.ppViewNormal;

                switch (type)
                {
                    case Type.Beam:
                        CreateBeamAgenda(slideTracker.SelectedSlides);
                        break;
                    case Type.Bullet:
                        CreateBulletAgenda();
                        break;
                    case Type.Visual:
                        CreateVisualAgenda();
                        break;
                }

                PowerPointPresentation.Current.AddAckSlide();
                SelectOriginalSlide(slideTracker.UserCurrentSlide, PowerPointPresentation.Current.FirstSlide);
            }
            finally
            {
                if (dialogOpen)
                {
                    DisposeLoadingDialog();
                }
                curWindow.ViewType = oldViewType;
            }
        }

        public static void RemoveAgenda()
        {
            var curWindow = Globals.ThisAddIn.Application.ActiveWindow;
            var oldViewType = curWindow.ViewType;

            try
            {
                var slideTracker = new SlideSelectionTracker(SelectedSlides, CurrentSlide);

                if (!AgendaPresent())
                {
                    MessageBox.Show(TextCollection.AgendaLabNoAgendaError);
                    return;
                }

                RemoveAllAgendaItems(slideTracker);

                SelectOriginalSlide(slideTracker.UserCurrentSlide, PowerPointPresentation.Current.FirstSlide);
            }
            finally
            {
                curWindow.ViewType = oldViewType;
            }
        }

        // PLACEHOLDER!
        public static void SynchroniseAgenda()
        {
            var slideTracker = new SlideSelectionTracker(SelectedSlides, CurrentSlide);

            var refSlide = FindReferenceSlide();
            var type = GetReferenceSlideType();

            BringToFront(refSlide);

            slideTracker.DeleteAcknowledgementSlideAndTrack();

            switch (type)
            {
                case Type.Beam:
                    break;
                case Type.Bullet:
                    SyncBulletAgenda(refSlide);
                    break;
                case Type.Visual:
                    break;
            }
        }
        #endregion


        #region FUNCTIONS

        /// <summary>
        /// Assumption: no reference slide exists
        /// </summary>
        private static void CreateBulletAgenda()
        {
            var refSlide = CreateBulletReferenceSlide();

            // here we invoke sync logic, since it's the same behavior as sync
            SyncBulletAgenda(refSlide);
        }


        private static void SelectOriginalSlide(PowerPointSlide originalSlide, PowerPointSlide fallbackToSlide)
        {
            if (originalSlide != null)
            {
                originalSlide.GetNativeSlide().Select();
                return;
            }
            if (fallbackToSlide != null)
            {
                fallbackToSlide.GetNativeSlide().Select();
            }
        }


        private static void RemoveAllAgendaItems(SlideSelectionTracker slideTracker = null)
        {
            if (slideTracker == null) slideTracker = SlideSelectionTracker.CreateInactiveTracker();

            PowerPointPresentation.Current.Slides.Where(AgendaSlide.IsAnyAgendaSlide)
                                                .ToList()
                                                .ForEach(slideTracker.DeleteSlideAndTrack);

            RemoveBeamAgendaFromSlides(PowerPointPresentation.Current.Slides);
        }

        private static void RemoveBeamAgendaFromSlides(IEnumerable<PowerPointSlide> candidates)
        {
            candidates = candidates.Where(AgendaSlide.IsNotReferenceslide);
            foreach (var candidate in candidates)
            {
                var beamShape = FindBeamShape(candidate);

                if (beamShape != null)
                {
                    beamShape.Delete();
                }
            }
        }

        private static void BringToFront(PowerPointSlide slide)
        {
            slide.MoveTo(1);
        }

        private static void SyncBulletAgenda(PowerPointSlide refSlide)
        {
            var sections = Sections;

            ScrambleSlideSectionNames();
            foreach (var section in sections)
            {
                RebuildBulletSectionSlides(section);
                SyncBulletSectionSlides(refSlide, section);
            }
        }

        /// <summary>
        /// Scrambles the slide section names to avoid duplicate names later on, which can crash powerpoint.
        /// Use this just before reassigning the slide section names! Don't keep the slide names this way!
        /// </summary>
        private static void ScrambleSlideSectionNames()
        {
            var slides = PowerPointPresentation.Current.Slides;
            slides.Where(slide => AgendaSlide.IsAnyAgendaSlide(slide) && AgendaSlide.IsNotReferenceslide(slide))
                    .ToList()
                    .ForEach(AgendaSlide.AssignUniqueSectionName);
        }

        /// <summary>
        /// 
        /// </summary>
        private static void RebuildBulletSectionSlides(AgendaSection section)
        {
            RebuildSectionUsingTemplate(section, new BulletAgendaTemplate());
        }

        private static void SyncBulletSectionSlides(PowerPointSlide refSlide, AgendaSection section)
        {
            //throw new NotImplementedException();
        }


        private static void SyncVisualAgenda()
        {
            var sections = Sections;
            var sectionMappings = GetSectionMappings();
            RemapVisualAgendaImages(sectionMappings);

            ScrambleSlideSectionNames();
            foreach (var section in sections)
            {
                RebuildVisualSectionSlides(section);
                SyncVisualAgendaSectionSlides(section);
            }
        }

        private static void SyncVisualAgendaSectionSlides(AgendaSection section)
        {
            throw new NotImplementedException();
        }

        private static void RebuildVisualSectionSlides(AgendaSection section)
        {
            throw new NotImplementedException();
        }

        private static void RemapVisualAgendaImages(Dictionary<int, int> sectionMappings)
        {
            throw new NotImplementedException();
        }

        /// <summary>
        /// Section mappings are of the form [new section index -> old section index]
        /// </summary>
        /// <returns></returns>
        private static Dictionary<int,int> GetSectionMappings()
        {
            throw new NotImplementedException();
        }

        private static void SyncBeamAgenda()
        {
            SyncBeamAgendaSlides();
        }

        private static void SyncBeamAgendaSlides()
        {
            throw new NotImplementedException();
        }

        private static bool DisplayLoadingDialog(string title, string content)
        {
            _loadDialog = new LoadingDialog(title, content);
            _loadDialog.Show();
            _loadDialog.Refresh();
            return true;
        }

        private static void DisposeLoadingDialog()
        {
            _loadDialog.Dispose();
        }

        #endregion

        #region Sync Functions

        public static SyncFunction SyncVisualAgendaSlide = (refSlide, sections, currentSection, targetSlide) =>
        {
            throw new NotImplementedException();
        };

        public static SyncFunction SyncBulletAgendaSlide = (refSlide, sections, currentSection, targetSlide) =>
        {
            throw new NotImplementedException();
        };

        #endregion


        #region Conditions on current state

        private static bool ValidSections()
        {
            var sections = PowerPointPresentation.Current.Sections;

            if (sections.Count == 0)
            {
                MessageBox.Show(TextCollection.AgendaLabNoSectionError);
                return false;
            }

            if (sections.Count == 1)
            {
                MessageBox.Show(TextCollection.AgendaLabSingleSectionError);
                return false;
            }

            if (PowerPointPresentation.Current.HasEmptySection)
            {
                MessageBox.Show(TextCollection.AgendaLabEmptySectionError);
                return false;
            }

            return true;
        }
        private static bool IsReferenceSlidePresent()
        {
            return FindReferenceSlide() != null;
        }

        private static bool AgendaPresent()
        {
            return FindAllAgendaSlides().Count > 0 || FindSlidesWithBeam().Count > 0;
        }

        #endregion



        #region UTILITY


        /// <summary>
        /// If user has slides selected, return a list of the selected slides.
        /// If user is not currently selecting slides, returns an empty list.
        /// </summary>
        private static List<PowerPointSlide> SelectedSlides
        {
            get
            {
                if (PowerPointCurrentPresentationInfo.CurrentSelection.Type != PpSelectionType.ppSelectionSlides)
                    return new List<PowerPointSlide>();
                return PowerPointPresentation.Current.SelectedSlides;
            }
        }


        /// <summary>
        /// Return the slide the user is currently on.
        /// </summary>
        private static PowerPointSlide CurrentSlide
        {
            get
            {
                var slides = PowerPointPresentation.Current.SelectedSlides;
                if (slides.Count == 0) return null;
                return slides[0];
            }
        }

        private static void RemoveRemovedSlides(ref List<PowerPointSlide> selection, ref PowerPointSlide slideSelectedByUser, List<string> removedSlideNames)
        {
            selection = selection.Where(slide => !removedSlideNames.Contains(slide.Name)).ToList();
            slideSelectedByUser = removedSlideNames.Contains(slideSelectedByUser.Name) ? null : slideSelectedByUser;
        }


        private static PowerPointSlide FindReferenceSlide()
        {
            var slides = PowerPointPresentation.Current.Slides;
            return slides.FirstOrDefault(AgendaSlide.IsReferenceslide);
        }

        private static List<PowerPointSlide> FindAllAgendaSlides()
        {
            var slides = PowerPointPresentation.Current.Slides;
            return slides.Where(AgendaSlide.IsAnyAgendaSlide).ToList();
        }

        private static List<PowerPointSlide> FindSlidesWithBeam()
        {
            var slides = PowerPointPresentation.Current.Slides;
            return slides.Where(HasBeamShape).ToList();
        }

        private static Shape FindBeamShape(PowerPointSlide slide)
        {
            return slide.Shapes.Cast<Shape>().FirstOrDefault(AgendaShape.IsBeamShape);
        }

        private static bool HasBeamShape(PowerPointSlide slide)
        {
            return slide.Shapes.Cast<Shape>().Any(AgendaShape.IsBeamShape);
        }

        private static Type GetReferenceSlideType()
        {
            var referenceSlide = FindReferenceSlide();
            if (referenceSlide == null) return Type.None;
            return AgendaSlide.Decode(referenceSlide).AgendaType;
        }

        /// <summary>
        /// 1-indexed.
        /// </summary>
        private static List<AgendaSection> Sections
        {
            get
            {
                // TODO: This is a zip-with-index code. I can rephrase this few lines better (more functional)
                var sectionNames = PowerPointPresentation.Current.Sections;
                var sections = new List<AgendaSection>();
                for (var i = 0; i < sectionNames.Count; ++i)
                {
                    int index = i + 1;
                    sections.Add(new AgendaSection(sectionNames[i], index));
                }
                return sections;
            }
        }


        /// <summary>
        /// Returns -1 when old section index is not found.
        /// </summary>
        private static int GetOldSectionIndex(AgendaSection section)
        {
            var sectionSlides = GetSectionSlides(section);
            foreach (var slide in sectionSlides)
            {
                var agendaSlide = AgendaSlide.Decode(slide);
                if (agendaSlide != null)
                {
                    return agendaSlide.Section.Index;
                }
            }
            return -1;
        }

        private static int NumberOfSections
        {
            get { return PowerPointPresentation.Current.Sections.Count; }
        }

        private static int FindSectionStart(AgendaSection section)
        {
            throw new NotImplementedException();
        }

        private static int FindSectionEnd(AgendaSection section)
        {
            throw new NotImplementedException();
        }

        private static AgendaSection FindSlideSection(PowerPointSlide slide)
        {
            throw new NotImplementedException();
        }


        /// <summary>
        /// the function will return the start agenda slide if the first slide of the requested
        /// section is an agenda slide, else it will return null. It also modify the name of the
        /// start slide to adapt the section's name change.
        /// 
        /// if it's beam type or none type, return the slide immediately. None type should be
        /// used if the user wants to return the first slide of each section regardless if
        /// it's an agenda slide.
        /// </summary>
        private static PowerPointSlide FindSectionStartSlide(AgendaSection section)
        {
            throw new NotImplementedException();
        }

        private static PowerPointSlide FindSectionEndSlide(AgendaSection section)
        {
            throw new NotImplementedException();
        }

        /// <summary>
        /// 1-indexed.
        /// </summary>
        private static int SectionFirstSlideIndex(AgendaSection section)
        {
            var sectionProperties = PowerPointPresentation.Current.SectionProperties;
            return sectionProperties.FirstSlide(section.Index);
        }

        /// <summary>
        /// 1-indexed
        /// </summary>
        private static int SectionLastSlideIndex(AgendaSection section)
        {
            var sectionProperties = PowerPointPresentation.Current.SectionProperties;
            int lastSlideIndex = PowerPointPresentation.Current.SlideCount;

            if (!IsLastSection(section))
            {
                lastSlideIndex = sectionProperties.FirstSlide(section.Index + 1) - 1;
            }

            return lastSlideIndex;
        }

        /// <summary>
        /// 0-indexed.
        /// </summary>
        private static List<PowerPointSlide> GetSectionSlides(AgendaSection section, bool excludeReferenceSlide = false)
        {
            var slides = PowerPointPresentation.Current.Slides;

            int firstSlideIndex = SectionFirstSlideIndex(section);
            int lastSlideIndex = SectionLastSlideIndex(section);

            return slides.GetRange(firstSlideIndex - 1, lastSlideIndex - firstSlideIndex + 1);
        }

        private static bool IsLastSection(AgendaSection section)
        {
            return section.Index == NumberOfSections;
        }

        #endregion



        #region NEED REFACTORING

        /// <summary>
        /// Generates the beam agenda on the target slides. Skips over the Reference (Template) slide if included in targetSlides.
        /// Generates the Reference slide if it does not already exist.
        /// Leave the targetSlides field blank (=null) to generate the beam agenda over all slides (other than the first section).
        /// </summary>
        private static void CreateBeamAgenda(IEnumerable<PowerPointSlide> targetSlides = null)
        {
            var sections = Sections;

            List<PowerPointSlide> slides;
            if (targetSlides != null)
            {
                slides = targetSlides.Where(AgendaSlide.IsNotReferenceslide).ToList();
            }
            else
            {
                var firstSectionIndex = FindSectionStart(sections[0]);
                slides = PowerPointPresentation.Current.Slides
                    .Where(slide => slide.Index >= firstSectionIndex && AgendaSlide.IsNotReferenceslide(slide))
                    .ToList();
            }

            if (slides.Count < 1) return;

            var refSlide = FindReferenceSlide();
            bool generateNewReferenceSlide;
            if (AgendaSlide.IsReferenceslide(refSlide))
            {
                var beamShape = FindBeamShape(refSlide);
                if (beamShape != null)
                {
                    generateNewReferenceSlide = false;
                    beamShape.Copy();
                }
                else
                {
                    // reference slide doesn't have a beam shape. weird. so we delete and recreate.
                    generateNewReferenceSlide = true;
                    refSlide.Delete();
                }
            }
            else
            {
                // can't find a reference slide.
                generateNewReferenceSlide = true;
            }

            if (generateNewReferenceSlide)
            {
                // if we do not have legacy template, create a new refslide 
                refSlide = CreateBeamReferenceSlide();

                PrepareBeamAgendaShapes(sections, refSlide);
                AddAgendaSlideBeamType(sections[0], refSlide);
                refSlide.BringIndicatorToFront();
            }

            // The beam shape is now stored in the clipboard to be pasted on each of the slides.
            foreach (var slide in slides)
            {
                AddAgendaSlideBeamType(FindSlideSection(slide), slide);
            }
        }


        private static void SyncAgendaBullet(List<AgendaSection> sections, PowerPointSlide refSlide, bool enableButtons = false)
        {
            AdjustBulletTemplateContent(sections);
            
            var referenceContentShape = refSlide.GetShape(AgendaShape.WithPurpose(ShapePurpose.ContentShape));
            var bulletFormats = BulletFormats.ExtractFormats(referenceContentShape);

            foreach (var section in sections)
            {
                var start = FindSectionStartSlide(section);
                var end = FindSectionEndSlide(section);

                SyncSingleAgendaBullet(bulletFormats, start, section, false, section.Index);
                SyncSingleAgendaBullet(bulletFormats, end, section, true, section.Index);
            }

            if (enableButtons)
            {
                throw new NotImplementedException();
                // TODO: RENABLE THIS
                /*foreach (var section in sections)
                {
                    AddLinkBulletAgenda(FindSectionStartSlide(section, Type.Bullet));
                    AddLinkBulletAgenda(FindSectionEndSlide(section, Type.Bullet));
                }*/
            }
        }

        private static void SyncSingleAgendaBullet(BulletFormats bulletFormats, PowerPointSlide candidate,
                                                   AgendaSection section, bool isEnd, int focusIndex)
        {
            throw new NotImplementedException();
            // if this is a new section, we need to generate a new agenda slide, else we need to check
            // if the slide's content is outdated. If so, we need to update the content and reformat it
            // according to the refslide.
            /*var refContentHolder = refSlide.GetShapeWithName(PptLabsAgendaContentShapeName)[0];
            var contentHolder = (candidate ?? (candidate = AddAgendaSlideBulletType(section, isEnd, refSlide))).GetShapeWithName(PptLabsAgendaContentShapeName)[0];

            if (_agendaOutdated)
            {
                contentHolder.TextFrame.TextRange.Text = _agendaText;
            }

            // after syncing the content, we need to take care of the general slide settings
            SyncSingleAgendaGeneral(refSlide, candidate);

            // then we sync the content holder without modifying the content
            Utils.Graphics.SyncShape(refContentHolder, contentHolder,
                                     pickupTextContent: false, pickupTextFormat: false);

            // finally recolor the bullets
            ReformatTextRange(contentHolder.TextFrame2.TextRange, focusIndex);*/
        }

        private static void AdjustBulletTemplateContent(List<AgendaSection> sections)
        {
            var refSlide = FindReferenceSlide();
            throw new NotImplementedException();

            // post process bullet points
            /*var contentHolder = refSlide.GetShape(AgendaShape.WithPurpose(ShapePurpose.ContentShape));
            var textRange = contentHolder.TextFrame2.TextRange;

            while (textRange.Paragraphs.Count < totalSection)
            {
                textRange.InsertAfter("\r ");
            }

            while (textRange.Paragraphs.Count > 3 && textRange.Paragraphs.Count > totalSection)
            {
                textRange.Paragraphs[textRange.Paragraphs.Count].Delete();
            }

            for (var i = 4; i <= textRange.Paragraphs.Count; i++)
            {
                textRange.Paragraphs[i].ParagraphFormat.Bullet.Type = MsoBulletType.msoBulletNone;
            }*/
        }


        private static void CreateVisualAgenda()
        {
            throw new NotImplementedException();
            /*
            var sections = Sections;

            PrepareVisualAgendaSlideCapture(sections);

            var refSlide = CreateVisualReferenceSlide();

            SyncAgendaVisual(sections, refSlide);*/
        }

        private static PowerPointSlide CreateBeamReferenceSlide()
        {
            var refSlide = PowerPointSlide.FromSlideFactory(PowerPointPresentation.Current
                                                            .Presentation
                                                            .Slides
                                                            .Add(1, PpSlideLayout.ppLayoutBlank));
            AgendaSlide.SetAsReferenceSlideName(refSlide, Type.Beam);
            refSlide.AddTemplateSlideMarker();
            refSlide.Hidden = true;

            return refSlide;
        }

        private static PowerPointSlide CreateBulletReferenceSlide()
        {
            var refSlide = PowerPointSlide.FromSlideFactory(PowerPointPresentation.Current
                                                            .Presentation
                                                            .Slides
                                                            .Add(1, PpSlideLayout.ppLayoutText));

            var titleShape = refSlide.Shapes.Placeholders[1];
            var contentShape = refSlide.Shapes.Placeholders[2];
            AgendaShape.SetShapeName(titleShape, ShapePurpose.TitleShape, AgendaSection.None);
            AgendaShape.SetShapeName(contentShape, ShapePurpose.ContentShape, AgendaSection.None);

            Graphics.SetText(titleShape, TextCollection.AgendaLabBulletTitleContent);
            Graphics.SetText(contentShape, TextCollection.AgendaLabBulletVisitedContent,
                                            TextCollection.AgendaLabBulletHighlightedContent,
                                            TextCollection.AgendaLabBulletUnvisitedContent);

            var paragraphs = Graphics.GetParagraphs(contentShape);
            paragraphs[0].Font.Fill.ForeColor.RGB = Graphics.ConvertColorToRgb(Color.Gray);
            paragraphs[1].Font.Fill.ForeColor.RGB = Graphics.ConvertColorToRgb(Color.Red);
            paragraphs[2].Font.Fill.ForeColor.RGB = Graphics.ConvertColorToRgb(Color.Black);

            AgendaSlide.SetAsReferenceSlideName(refSlide, Type.Bullet);
            refSlide.AddTemplateSlideMarker();
            refSlide.Hidden = true;
            refSlide.DeleteIndicator();

            return refSlide;
        }



        private static PowerPointSlide CreateVisualReferenceSlide()
        {
            throw new NotImplementedException();
        }


        private static void PrepareBeamAgendaShapes(List<AgendaSection> sections, PowerPointSlide refSlide)
        {
            throw new NotImplementedException();
        }

        private static void AddAgendaSlideBeamType(AgendaSection section, PowerPointSlide slide)
        {
            throw new NotImplementedException();
        }




        private static string CreateInDocHyperLink(PowerPointSlide slide)
        {
            throw new NotImplementedException();
        }


        #endregion

    }

}
