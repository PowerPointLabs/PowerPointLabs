using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using PowerPointLabs.Models;
using PowerPointLabs.Views;
using Shape = Microsoft.Office.Interop.PowerPoint.Shape;

namespace PowerPointLabs
{
    internal static class AgendaLab
    {
        private const string PptLabsAgendaSlideNameFormat = "PptLabs{0}Agenda{1}Slide {2}";
        private const string PptLabsAgendaTitleShapeName = "PptLabsAgendaTitle";
        private const string PptLabsAgendaContentShapeName = "PptLabsAgendaContent";
        private const string PptLabsAgendaSlideTypeSearchPattern = @"PptLabs(\w+)Agenda(?:Start|End)?Slide";
        private const string PptLabsAgendaVisualSectionName = "PptLabsAgendaVisualSection";
        private const string PptLabsAgendaVisualItemPrefix = "PptLabsAgendaVisualItem";
        private const string PptLabsAgendaBeamBackgroundName = "PptLabsAgendaBeamBackground";
        private const string PptLabsAgendaBeamShapeName = "PptLabsAgendaBeamShape";
        private const string PptLabsAgendaBeamHighlight = "PptLabsAgendaBeamHighlight";

        private const float VisualAgendaItemMargin = 0.05f;

        private static LoadingDialog _loadDialog = new LoadingDialog();

        private static readonly Regex AgendaSlideSearchPattern = new Regex(PptLabsAgendaSlideTypeSearchPattern);

        private static readonly string SlideCapturePath = Path.Combine(Path.GetTempPath(), "PowerPointLabs Temp");

        private static string _agendaText;

        private static bool _agendaOutdated;

        private static Color _bulletDefaultColor = Color.Black;
        private static Color _bulletHighlightColor = Color.Red;
        private static Color _bulletDimColor = Color.Gray;

        # region Enum
        public enum Type
        {
            None,
            Bullet,
            Beam,
            Visual
        };

        public enum Direction
        {
            Top,
            Left,
            Bottom,
            Right
        };
        # endregion

        /*********************************************************
         * Note:
         * The implementation of CurrentType is O(n),
         * bear this in mind when design functions. E.g. 
         * FindSectionEndSlide is totally fine with only the first
         * argument and check CurrentType in the function.
         * But taking type as the second argument allows the 
         * developer to get CurrentType before the function.
         * This is useful when CurrentType is frequently used
         * in a certain region, which reduces the overhead.
         * 
         * See SynchronizeAgenda() for more usage details.
         *********************************************************/
        # region Properties
        public static Type CurrentType
        {
            get
            {
                if (PowerPointPresentation.Current.SectionProperties.Count == 0)
                {
                    return Type.None;
                }

                var slides = PowerPointPresentation.Current.Slides;

                foreach (var slide in slides)
                {
                    if (AgendaSlideSearchPattern.IsMatch(slide.Name))
                    {
                        var type = AgendaSlideSearchPattern.Match(slide.Name).Groups[1].Value;

                        return (Type)Enum.Parse(typeof(Type), type);
                    }

                    // here we try to find the first occurance, or potential occurance of beam type agenda.
                    // A potential occurance of beam type agenda is a slide that contains a group shape which
                    // contains a shape with PptLabsAgendaBeamBackgroundName as name. We need to do this since
                    // a user may want to ungroup a beam agenda and change the item's format, after he re-group
                    // the shape, the original name will be changed, and thus need to be renamed for easier
                    // recognition.
                    var beamShape = FindBeamShape(slide);

                    if (beamShape != null)
                    {
                        return Type.Beam;
                    }
                }

                return Type.None;
            }
        }

        public static Direction BeamDirection { get; set; }
        # endregion

        # region API
        public static void BulletAgendaSettings()
        {
            PickupColorSettings();

            var settingDialog = new BulletAgendaSettingsDialog(_bulletHighlightColor,
                                                               _bulletDimColor,
                                                               _bulletDefaultColor);

            settingDialog.SettingsHandler += UpdateColorScheme;
            settingDialog.ShowDialog();
        }

        public static void GenerateAgenda(Type type)
        {
            // agenda exists in current presentation
            if (CurrentType != Type.None)
            {
                var confirm = MessageBox.Show(TextCollection.AgendaLabAgendaExistError,
                                              TextCollection.AgendaLabAgendaExistErrorCaption,
                                              MessageBoxButtons.OKCancel);

                if (confirm != DialogResult.OK) return;

                RemoveAgenda();
            }

            // validate section information
            if (!SectionValidation()) return;

            var selectedSlides = PowerPointCurrentPresentationInfo.SelectedSlides.ToList();
            var slides = PowerPointPresentation.Current.Slides;

            if (type == Type.Beam && selectedSlides.Count == 0)
            {
                MessageBox.Show(TextCollection.AgendaLabNoSelectionError);
                return;
            }

            var sections = PowerPointPresentation.Current.Sections.Skip(1).ToList();

            _loadDialog = new LoadingDialog(TextCollection.AgendaLabLoadingDialogTitle,
                                                TextCollection.AgendaLabLoadingDialogContent);
            _loadDialog.Show();
            _loadDialog.Refresh();

            var curWindow = Globals.ThisAddIn.Application.ActiveWindow;
            var oldViewType = curWindow.ViewType;

            curWindow.ViewType = PpViewType.ppViewNormal;

            switch (type)
            {
                case Type.Beam:
                    GenerateBeamAgenda(sections, selectedSlides);
                    break;
                case Type.Bullet:
                    GenerateBulletAgenda(sections);
                    break;
                case Type.Visual:
                    GenerateVisualAgenda(sections);
                    break;
            }

            PowerPointPresentation.Current.AddAckSlide();

            curWindow.ViewType = oldViewType;
            SelectOriginalSlide(selectedSlides.Count > 0 ? selectedSlides[0] : slides[0], slides[0]);

            _loadDialog.Dispose();
        }

        public static void RemoveAgenda()
        {
            var type = CurrentType;

            if (type == Type.None)
            {
                MessageBox.Show(TextCollection.AgendaLabNoAgendaError);
                return;
            }

            var selectedSlides = PowerPointCurrentPresentationInfo.SelectedSlides.ToList();
            var slides = PowerPointPresentation.Current.Slides;

            switch (type)
            {
                case Type.Beam:
                    RemoveBeamAgenda(selectedSlides);
                    break;
                case Type.Bullet:
                    RemoveBulletAgenda();
                    break;
                case Type.Visual:
                    RemoveVisualAgenda();
                    break;
            }

            PowerPointPresentation.Current.RemoveAckSlide();

            SelectOriginalSlide(selectedSlides.Count > 0 ? selectedSlides[0] : slides[0], slides[0]);
        }

        /***********************************************************************************
         * Basically, sync will not only sync the format across all agenda slides, but
         * the content will also be adjusted to fit the current context. To achieve this,
         * we simply remove all slides and regenerate the slides before we proceed.
         * However, there are some details that should be noted:
         * 1. Reference finding mechanism is different for Beam style;
         * 2. After foudn the reference slide, we need to cut it to the clipboard;
         * 3. Before we remove all agenda, we should pick up color settings in case we are
         *    dealing with bullet agenda;
         * 4. Before we re-generate the agenda, we should paste the reference slide somewhere.
         *    In our implementation, we chose to duplicate the first slide in the first section
         *    then paste the reference slide in between. This will prevent generating incorrect
         *    transition when visual style is used.
         * The procedure stated above is encapsulated in PrepareSyncAgenda().
         *************************************************************************************/
        public static void SynchronizeAgenda()
        {
            var type = CurrentType;

            if (type == Type.None)
            {
                MessageBox.Show(TextCollection.AgendaLabNoAgendaError);
                return;
            }

            if (!SectionValidation()) return;

            var selectedSlides = PowerPointCurrentPresentationInfo.SelectedSlides.ToList();

            if (type == Type.Beam && selectedSlides.Count == 0)
            {
                MessageBox.Show(TextCollection.AgendaLabNoSelectionError);
                return;
            }

            _loadDialog = new LoadingDialog("Synchronizing...", "Agenda is getting synchronized, please wait...");
            _loadDialog.Show();
            _loadDialog.Refresh();

            var curWindow = Globals.ThisAddIn.Application.ActiveWindow;
            var oldViewType = curWindow.ViewType;

            curWindow.ViewType = PpViewType.ppViewNormal;

            // find the agenda for the first section as reference
            var currentPresentation = PowerPointPresentation.Current;
            var sections = currentPresentation.Sections.Where(section =>
                                                              section != PptLabsAgendaVisualSectionName).Skip(1).ToList();

            var refSlide = FindReferenceSlide(type);

            // refSlide will be copied and pasted to the beginning of the presentation as a
            // format reference, and all agenda slides will be deleted and regenerated to take
            // sections change into account. It needs to be deleted after sync has been done.
            PrepareSync(type, ref refSlide);

            // regenerate slides and sync accordingly
            switch (type)
            {
                case Type.Beam:
                    GenerateBeamAgenda(sections, selectedSlides);
                    SyncAgendaBeam(refSlide, selectedSlides);
                    break;
                case Type.Bullet:
                    GenerateBulletAgenda(sections);
                    SyncAgendaBullet(sections, refSlide);
                    break;
                case Type.Visual:
                    GenerateVisualAgenda(sections);
                    SyncAgendaVisual(sections, refSlide);
                    break;
            }

            refSlide.Delete();

            curWindow.ViewType = oldViewType;
            SelectOriginalSlide(selectedSlides[0], PowerPointPresentation.Current.Slides[0]);
            _loadDialog.Dispose();
        }
        # endregion

        # region Helper Functions
        private static void AddAgendaSlideBeamType(string section, PowerPointSlide slide)
        {
            var beams = slide.Shapes.Paste()[1].Ungroup();

            var currentTextBox = beams.Cast<Shape>().FirstOrDefault(box => box.TextFrame
                                                                              .TextRange
                                                                              .Text == section);

            if (currentTextBox != null)
            {
                currentTextBox.TextFrame.TextRange.Font.Color.RGB = Utils.Graphics.ConvertColorToRgb(Color.Yellow);
                currentTextBox.Name += " " + PptLabsAgendaBeamHighlight;
            }

            beams.Group().Name = PptLabsAgendaBeamShapeName;
        }

        private static void AddAgendaSlideBulletType(string section, bool isEnd)
        {
            var sectionIndex = FindSectionIndex(section);
            var sectionEndIndex = FindSectionEnd(section);

            var slide =
                PowerPointSlide.FromSlideFactory(PowerPointPresentation.Current
                                                                       .Presentation
                                                                       .Slides
                                                                       .Add(isEnd ? sectionEndIndex + 1 : 1,
                                                                            PpSlideLayout.ppLayoutText));
            if (!isEnd)
            {
                slide.GetNativeSlide().MoveToSectionStart(sectionIndex);
            }

            slide.Name = string.Format(PptLabsAgendaSlideNameFormat, Type.Bullet,
                                          isEnd ? "End" : "Start", section);
            slide.Transition.EntryEffect = PpEntryEffect.ppEffectFadeSmoothly;
            slide.Transition.Duration = 0.25f;

            slide.Shapes.Placeholders[1].Name = PptLabsAgendaTitleShapeName;
            slide.Shapes.Placeholders[2].Name = PptLabsAgendaContentShapeName;

            // set title
            slide.Shapes.Placeholders[1].TextFrame.TextRange.Text = "Agenda";

            // set agenda content
            var contentPlaceHolder = slide.Shapes.Placeholders[2];
            var textRange = contentPlaceHolder.TextFrame.TextRange;
            var focusColor = _bulletHighlightColor;

            // since section index is 1-based, focus section index should be substracted by 1
            RecolorTextRange(textRange, sectionIndex - 1, focusColor);
        }

        private static void AddAgendaSlideVisualType(List<string> sections)
        {
            var currentPresentation = PowerPointPresentation.Current.Presentation;
            var sectionProperties = PowerPointPresentation.Current.SectionProperties;

            for (var i = 0; i <= sections.Count; i++)
            {
                var sectionName = i < sections.Count ? sections[i] : sections[i - 1];
                var prevSectionName = i == 0 ? string.Empty : sections[i - 1];

                // add a new section before the first section and rename to PptLabsAgendaVisualSectionName
                var index = i < sections.Count ? FindSectionStart(sectionName) : PowerPointPresentation.Current.SlideCount;
                var nativeSlide = i == 0 ? currentPresentation.Slides.Add(index, PpSlideLayout.ppLayoutTitleOnly) :
                                           currentPresentation.Slides.Paste(index)[1];
                var slide = PowerPointSlide.FromSlideFactory(nativeSlide);

                slide.Name = string.Format(PptLabsAgendaSlideNameFormat, Type.Visual,
                                           string.Empty, i < sections.Count ? sectionName : "EndOfAgenda");
                var newSectionIndex = sectionProperties.AddBeforeSlide(index, PptLabsAgendaVisualSectionName);

                // if we are in the first agenda section, generate slide shapes in the canvas area,
                // else we should generate step back effect
                if (i == 0)
                {
                    PrepareVisualAgendaSlideShapes(slide, sections);
                }

                if (i > 0)
                {
                    GenerateVisualAgendaSlideZoomOut(slide, prevSectionName, newSectionIndex);
                }

                if (i < sections.Count)
                {
                    GenerateVisualAgendaSlideZoomIn(slide, sectionName);

                    slide.Copy();
                }
            }
        }

        private static void AdjustBeamItemHorizontal(ref float lastLeft, ref float lastTop, ref float widest,
                                                     float delta, Shape item, Shape background)
        {
            if (lastLeft + delta > PowerPointPresentation.Current.SlideWidth)
            {
                lastLeft = 0;
                lastTop += item.Height;
            }

            item.Left = Math.Max(lastLeft, lastLeft + (delta - item.Width) / 2f);
            item.Top = lastTop;

            if (item.Width > widest)
            {
                widest = item.Width;
            }

            lastLeft += Math.Max(item.Width, delta);

            if (background.Height < lastTop + item.Height)
            {
                background.Height = lastTop + item.Height;
            }
        }

        private static void AdjustBeamItemVertical(ref float lastLeft, ref float lastTop, Shape item, Shape background)
        {
            lastTop += item.Height;

            if (lastTop >= PowerPointPresentation.Current.SlideHeight)
            {
                lastTop = 0;
                lastLeft += item.Width;

                item.Top = 0;
                item.Left = lastLeft;
            }

            if (background.Width < lastLeft + item.Width)
            {
                background.Width = lastLeft + item.Width;
            }
        }

        private static void CheckAgendaUpdate(Type type, PowerPointSlide refSlide, string refSection)
        {
            switch (type)
            {
                case Type.Beam:
                    CheckBeamUpdate(refSlide, refSection);
                    break;
                case Type.Visual:
                    CheckVisualUpdate(refSlide);
                    break;
            }
        }

        private static void CheckBeamUpdate(PowerPointSlide refSlide, string refSection)
        {
            var beamShape = refSlide.GetShapeWithName(PptLabsAgendaBeamShapeName)[0];

            // check if the section names have changed
            var sections = PowerPointPresentation.Current.Sections.Skip(1).OrderBy(x => x).ToList();
            var beamItems = beamShape.GroupItems.Cast<Shape>()
                                                .Select(shape => shape.Name)
                                                .Where(name => name != PptLabsAgendaBeamBackgroundName)
                                                .OrderBy(x => x).ToList();

            for (var i = 0; i < beamItems.Count; i ++)
            {
                if (beamItems[i].EndsWith(PptLabsAgendaBeamHighlight))
                {
                    var totalLen = beamItems[i].Length;
                    var suffixLen = PptLabsAgendaBeamHighlight.Length;
                    beamItems[i] = beamItems[i].Remove(totalLen - suffixLen - 1, suffixLen + 1);
                }
            }

            _agendaOutdated = !sections.SequenceEqual(beamItems);

            refSlide.Name = refSection;
        }

        private static void CheckVisualUpdate(PowerPointSlide refSlide)
        {
            var visualItems =
                refSlide.GetShapesWithPrefix(PptLabsAgendaVisualItemPrefix)
                        .Select(shape => shape.Name.Substring(PptLabsAgendaVisualItemPrefix.Length))
                        .OrderBy(x => x);
            var sections = PowerPointPresentation.Current
                                                 .Sections
                                                 .Skip(1)
                                                 .Where(section => section != PptLabsAgendaVisualSectionName)
                                                 .OrderBy(x => x).ToList();

            _agendaOutdated = !sections.SequenceEqual(visualItems);
        }

        private static string CreateInDocHyperLink(PowerPointSlide slide)
        {
            return slide.ID + "," + slide.Index + "," + slide.Name;
        }

        private static Shape FindBeamHighlight(List<Shape> beamItems)
        {
            return beamItems.FirstOrDefault(shape => shape.Name.EndsWith(PptLabsAgendaBeamHighlight));
        }

        private static Shape FindBeamNormal(List<Shape> beamItems)
        {
            return beamItems.FirstOrDefault(shape => !shape.Name.EndsWith(PptLabsAgendaBeamHighlight) &&
                                                     shape.Name != PptLabsAgendaBeamBackgroundName);
        }

        private static Shape FindBeamShape(PowerPointSlide slide)
        {
            var grouped = slide.GetShapesWithTypeAndRule(MsoShapeType.msoGroup, new Regex(".+"));

            if (grouped == null || grouped.Count == 0)
            {
                return null;
            }

            for (var i = 0; i < grouped.Count; i++)
            {
                if (Utils.Graphics.IsCorrupted(grouped[i]))
                {
                    grouped[i] = Utils.Graphics.CorruptionCorrection(grouped[i], slide);
                }
            }

            var beamShape = grouped.FirstOrDefault(shape => shape.Name == PptLabsAgendaBeamShapeName) ??
                            grouped.FirstOrDefault(shape => shape.GroupItems
                                                                 .Cast<Shape>()
                                                                 .Any(subItem => subItem.Name == PptLabsAgendaBeamBackgroundName));

            if (beamShape != null &&
                beamShape.Name != PptLabsAgendaBeamShapeName)
            {
                beamShape.Name = PptLabsAgendaBeamShapeName;
            }

            return beamShape;
        }

        private static PowerPointSlide FindReferenceSlide(Type type)
        {
            var slides = PowerPointPresentation.Current.Slides;

            if (type == Type.Beam)
            {
                return slides.FirstOrDefault(slide => slide.GetShapeWithName(PptLabsAgendaBeamShapeName).Count != 0);
            }

            var generatedSlideName = string.Format("PptLabs{0}Agenda", type);

            return slides.FirstOrDefault(slide => slide.Name.Contains(generatedSlideName));
        }

        private static int FindSectionEnd(string section)
        {
            var sectionIndex = FindSectionIndex(section);

            return FindSectionEnd(sectionIndex);
        }

        private static int FindSectionEnd(int sectionIndex)
        {
            // here the sectionIndex is 1-based!
            var sectionProperties = PowerPointPresentation.Current.SectionProperties;

            return sectionProperties.FirstSlide(sectionIndex) + sectionProperties.SlidesCount(sectionIndex) - 1;
        }

        private static PowerPointSlide FindSectionEndSlide(string section, Type type)
        {
            if (type == Type.Beam)
            {
                return null;
            }

            var slideName = string.Format(PptLabsAgendaSlideNameFormat, type,
                                          type == Type.Visual ? string.Empty : "End", section);

            return PowerPointPresentation.Current.Slides.FirstOrDefault(slide => slide.Name == slideName);
        }

        private static int FindSectionStart(string section)
        {
            var sectionIndex = FindSectionIndex(section);

            return FindSectionStart(sectionIndex);
        }

        private static int FindSectionStart(int sectionIndex)
        {
            // here the sectionIndex is 1-based!
            var sectionProperties = PowerPointPresentation.Current.SectionProperties;

            return sectionProperties.FirstSlide(sectionIndex);
        }

        private static PowerPointSlide FindSectionStartSlide(string section, Type type)
        {
            var curPresentation = PowerPointPresentation.Current;
            var slides = curPresentation.Slides;

            if (type == Type.Beam)
            {
                var sectionProperties = curPresentation.SectionProperties;
                var sectionIndex = FindSectionIndex(section);

                return slides[sectionProperties.FirstSlide(sectionIndex) - 1];
            }

            var slideName = string.Format(PptLabsAgendaSlideNameFormat, type, 
                                          type == Type.Visual ? string.Empty : "Start", section);

            return slides.FirstOrDefault(slide => slide.Name == slideName);
        }

        private static int FindSectionIndex(string section)
        {
            // here the return value is 1-based!
            return PowerPointPresentation.Current.Sections.FindIndex(name => name == section) + 1;
        }

        private static string FindSlideSection(PowerPointSlide slide)
        {
            var sections = PowerPointPresentation.Current.Sections;

            for (var i = 0; i < sections.Count; i ++)
            {
                var sectionStart = FindSectionStart(sections[i]);
                
                if (sectionStart == slide.Index)
                {
                    return sections[i];
                }

                if (sectionStart > slide.Index)
                {
                    return sections[i - 1];
                }
            }

            return sections[sections.Count - 1];
        }

        private static void GenerateBeamAgenda(List<string> sections, List<PowerPointSlide> selectedSlides)
        {
            var firstSectionIndex = FindSectionStart(FindSectionIndex(sections[0]));
            var slides = selectedSlides.Where(slide => slide.Index >= firstSectionIndex).ToList();

            if (slides.Count < 1) return;

            PrepareBeamAgendaShapes(sections, slides[0]);

            foreach (var slide in slides)
            {
                AddAgendaSlideBeamType(FindSlideSection(slide), slide);
            }
        }

        private static void GenerateBulletAgenda(List<string> sections)
        {
            // need to use '\r' as paragraph indicator, not '\n'!
            // must end with '\r' to make the last line a paragraph!
            _agendaText = sections.Aggregate((current, next) => current + "\r" + next) + "\r";

            foreach (var section in sections)
            {
                AddAgendaSlideBulletType(section, false);
                AddAgendaSlideBulletType(section, true);
            }
        }

        private static void GenerateVisualAgenda(List<string> sections)
        {
            PrepareVisualAgendaSlideCapture(sections);

            AddAgendaSlideVisualType(sections);
        }

        private static void GenerateVisualAgendaSlideZoomIn(PowerPointSlide slide, string sectionName)
        {
            // get the shape that represent current slide
            var slideShape = slide.GetShapeWithName(PptLabsAgendaVisualItemPrefix + sectionName)[0];

            // add drill down effect and clean up current slide by deleting drill down
            // shape and recover original slide shape visibility
            AutoZoom.AddDrillDownAnimation(slideShape, slide);
            slide.GetShapesWithRule(new Regex("PPTZoomIn"))[0].Delete();
            slideShape.Visible = MsoTriState.msoTrue;
        }

        private static void GenerateVisualAgendaSlideZoomOut(PowerPointSlide slide, string sectionName, int sectionIndex)
        {
            // get the shape that represent previous slide, change the fillup picture to
            // the end of slide picture
            var sectionEndName = string.Format("{0} End.png", sectionName);
            var slideShape = slide.GetShapeWithName(PptLabsAgendaVisualItemPrefix + sectionName)[0];

            slideShape.Fill.UserPicture(Path.Combine(SlideCapturePath, sectionEndName));

            // add step back effect  and clean up current slide by deleting step back
            // shape and recover original slide shape visibility
            AutoZoom.AddStepBackAnimation(slideShape, slide);
            slide.GetShapesWithRule(new Regex("PPTZoomOut"))[0].Delete();
            slideShape.Visible = MsoTriState.msoTrue;

            var index = slide.Index;
            // move the step back slide to the section
            PowerPointPresentation.Current.Presentation.Slides[index - 1].MoveToSectionStart(sectionIndex);
        }

        private static void PickupColorFromSlide(PowerPointSlide slide)
        {
            var contentPlaceHolder = slide.GetShapeWithName(PptLabsAgendaContentShapeName)[0];
            var paragraphs = contentPlaceHolder.TextFrame2.TextRange
                                               .Paragraphs.Cast<TextRange2>()
                                               .Where(paragraph => paragraph.ParagraphFormat.IndentLevel == 1).ToList();

            _bulletHighlightColor = Utils.Graphics.ConvertRgbToColor(paragraphs[0].Font.Fill.ForeColor.RGB);
            
            var state = 0;

            for (var i = 1; i < paragraphs.Count; i ++ )
            {
                var paraColor = Utils.Graphics.ConvertRgbToColor(paragraphs[i].Font.Fill.ForeColor.RGB);
                
                if (state == 0)
                {
                    if (paraColor != _bulletHighlightColor)
                    {
                        _bulletDimColor = _bulletHighlightColor;
                        _bulletHighlightColor = paraColor;
                        state = 1;
                    }
                } else
                if (state == 1)
                {
                    if (paraColor == _bulletHighlightColor)
                    {
                        _bulletHighlightColor = _bulletDimColor;
                    }

                    _bulletDefaultColor = paraColor;
                    break;
                }
            }
        }

        private static void PickupColorSettings()
        {
            var type = CurrentType;

            if (type != Type.Bullet) return;

            var slides = PowerPointPresentation.Current.Slides;
            var sectionNameSearchPatternFormat = string.Format(PptLabsAgendaSlideNameFormat, "Bullet",
                                                               "(?:Start|End)", "(\\w+)");
            var sectionNameSearchPattern = new Regex(sectionNameSearchPatternFormat);
            var slideCandidates = slides.Where(slide => sectionNameSearchPattern.IsMatch(slide.Name));

            foreach (var slide in slideCandidates.Where(candidate => sectionNameSearchPattern.IsMatch(candidate.Name)))
            {
                PickupColorFromSlide(slide);
            }
        }

        private static Shape PrepareBeamAgendaBackground(PowerPointSlide slide, bool horizontal)
        {
            var background = slide.Shapes.AddShape(MsoAutoShapeType.msoShapeRectangle, 0, 0, 0, 0);
            background.Name = PptLabsAgendaBeamBackgroundName;
            background.Line.Visible = MsoTriState.msoFalse;
            background.Fill.ForeColor.RGB = Utils.Graphics.ConvertColorToRgb(Color.Black);

            if (horizontal)
            {
                background.Width = PowerPointPresentation.Current.SlideWidth;
            }
            else
            {
                background.Height = PowerPointPresentation.Current.SlideHeight;
            }

            return background;
        }

        private static Shape PrepareBeamAgendaBeamItem(PowerPointSlide slide, float lastLeft,
                                                       float lastTop, string section)
        {
            var textBox = slide.Shapes.AddTextbox(MsoTextOrientation.msoTextOrientationHorizontal,
                                                  lastLeft, lastTop, 0, 0);

            textBox.Name = section;
            textBox.TextFrame.AutoSize = PpAutoSize.ppAutoSizeShapeToFitText;
            textBox.TextFrame.WordWrap = MsoTriState.msoFalse;
            textBox.TextFrame.TextRange.Text = section;
            textBox.TextFrame.TextRange.Font.Color.RGB = Utils.Graphics.ConvertColorToRgb(Color.White);

            var mouseOnClickAction = textBox.ActionSettings[PpMouseActivation.ppMouseClick];

            mouseOnClickAction.Action = PpActionType.ppActionNamedSlideShow;
            mouseOnClickAction.Hyperlink.Address = null;
            mouseOnClickAction.Hyperlink.SubAddress = CreateInDocHyperLink(FindSectionStartSlide(section, Type.Beam));

            return textBox;
        }

        private static void PrepareBeamAgendaShapes(List<string> sections, PowerPointSlide refSlide)
        {
            var lastLeft = 0.0f;
            var lastTop = 0.0f;
            var slideWidth = PowerPointPresentation.Current.SlideWidth;
            var slideHeight = PowerPointPresentation.Current.SlideHeight;

            var horizontal = BeamDirection == Direction.Top || BeamDirection == Direction.Bottom;
            var background = PrepareBeamAgendaBackground(refSlide, horizontal);
            var widest = 0.0f;

            foreach (var section in sections)
            {
                var textBox = PrepareBeamAgendaBeamItem(refSlide, lastLeft, lastTop, section);

                if (horizontal)
                {
                    AdjustBeamItemHorizontal(ref lastLeft, ref lastTop, ref widest, 0, textBox, background);
                } else
                {
                    AdjustBeamItemVertical(ref lastLeft, ref lastTop, textBox, background);
                }
            }

            // for horizontal case, we need to evenly distribute every item
            if (horizontal)
            {
                background.Height = 0;
                lastLeft = 0;
                lastTop = 0;

                var delta = Math.Max(widest, slideWidth / sections.Count);

                foreach (var section in sections)
                {
                    var textBox = refSlide.GetShapeWithName(section)[0];

                    AdjustBeamItemHorizontal(ref lastLeft, ref lastTop, ref widest, delta, textBox, background);
                }
            }

            var copyRange = new List<string>();
            copyRange.AddRange(sections);
            copyRange.Add(PptLabsAgendaBeamBackgroundName);

            var group = refSlide.Shapes.Range(copyRange.ToArray()).Group();
            group.Name = PptLabsAgendaBeamShapeName;

            if (BeamDirection == Direction.Bottom)
            {
                group.Top = slideHeight - group.Height;
            } else
            if (BeamDirection == Direction.Right)
            {
                group.Left = slideWidth - group.Width;
            }

            group.Cut();
        }

        private static void PrepareSync(Type type, ref PowerPointSlide refSlide)
        {
            var refSection = FindSlideSection(refSlide);

            // pick up color setting before we remove the agenda when the type is bullet
            if (type == Type.Bullet)
            {
                PickupColorSettings();
            }
            
            refSlide.GetNativeSlide().Copy();
            var refDesign = refSlide.Design;
            refSlide = PowerPointSlide.FromSlideFactory(PowerPointPresentation.Current.Presentation.Slides.Paste(1)[1]);
            refSlide.Design = refDesign;

            CheckAgendaUpdate(type, refSlide, refSection);

            RemoveAgenda();
        }

        private static void PrepareVisualAgendaSlideCapture(IEnumerable<string> sections)
        {
            if (!Directory.Exists(SlideCapturePath))
            {
                Directory.CreateDirectory(SlideCapturePath);
            }

            var slides = PowerPointPresentation.Current.Slides;

            foreach (var section in sections)
            {
                var sectionStartSlide = slides[FindSectionStart(section) - 1];
                var sectionEndSlide = slides[FindSectionEnd(section) - 1];
                var animatedEndSlide = sectionEndSlide.Duplicate();

                foreach (var shape in animatedEndSlide.Shapes.Cast<Shape>().Where(animatedEndSlide.HasExitAnimation))
                {
                    shape.Delete();
                }

                animatedEndSlide.MoveMotionAnimation();

                var sectionStartName = string.Format("{0} Start.png", section);
                var sectionEndName = string.Format("{0} End.png", section);

                Utils.Graphics.ExportSlide(sectionStartSlide, Path.Combine(SlideCapturePath, sectionStartName));
                Utils.Graphics.ExportSlide(animatedEndSlide, Path.Combine(SlideCapturePath, sectionEndName));

                animatedEndSlide.Delete();
            }
        }

        private static void PrepareVisualAgendaSlideShapes(PowerPointSlide slide, List<string> sections)
        {
            var titleBar = slide.Shapes.Placeholders[1];

            titleBar.Name = PptLabsAgendaTitleShapeName;
            titleBar.TextFrame.TextRange.Text = "Agenda";

            var slideWidth = PowerPointPresentation.Current.SlideWidth;
            var slideHeight = PowerPointPresentation.Current.SlideHeight;
            var aspectRatio = slideWidth / slideHeight;
            var epsilon = slideHeight * 0.02f;

            var canvasLeft = titleBar.Left;
            var canvasTop = titleBar.Top + titleBar.Height + epsilon;
            var canvasWidth = titleBar.Width;
            var canvasHeight = canvasWidth / aspectRatio;

            var itemCount = sections.Count;
            var itemCanvasWidth = canvasWidth / itemCount;
            var itemWidth = itemCanvasWidth * (1 - 2 * VisualAgendaItemMargin);
            var itemHeight = itemWidth / aspectRatio;
            var itemTop = canvasTop + (canvasHeight - itemHeight) / 2;
            
            for (var i = 0; i < itemCount; i ++)
            {
                var itemLeft = canvasLeft + i*itemCanvasWidth + itemCanvasWidth * VisualAgendaItemMargin;

                var shape = slide.Shapes.AddShape(MsoAutoShapeType.msoShapeRectangle,
                                                  itemLeft, itemTop,
                                                  itemWidth, itemHeight);

                shape.Name = PptLabsAgendaVisualItemPrefix + sections[i];
                shape.Line.Visible = MsoTriState.msoFalse;
                shape.LockAspectRatio = MsoTriState.msoTrue;

                var slideCaptureName = string.Format("{0} Start.png", sections[i]);
                shape.Fill.UserPicture(Path.Combine(SlideCapturePath, slideCaptureName));
            }
        }

        private static void RecolorTextRange(TextRange textRange, int focusIndex, Color focusColor)
        {
            textRange.Font.Color.RGB = Utils.Graphics.ConvertColorToRgb(_bulletDefaultColor);

            textRange.Text = _agendaText;

            for (var i = 1; i < focusIndex; i++)
            {
                textRange.Paragraphs(i).Font.Color.RGB = Utils.Graphics.ConvertColorToRgb(_bulletDimColor);
            }

            textRange.Paragraphs(focusIndex).Font.Color.RGB = Utils.Graphics.ConvertColorToRgb(focusColor);
        }

        private static void RemoveBeamAgenda(List<PowerPointSlide> candidates)
        {
            foreach (var candidate in candidates)
            {
                var beamShape = FindBeamShape(candidate);

                if (beamShape != null)
                {
                    beamShape.Delete();
                }
            }
        }

        private static void RemoveBulletAgenda()
        {
            PowerPointPresentation.Current.RemoveSlide(AgendaSlideSearchPattern, true);
        }

        private static void RemoveVisualAgenda()
        {
            var curPresentation = PowerPointPresentation.Current;
            var sectionProperties = curPresentation.SectionProperties;
            var section = curPresentation.Sections;

            // to avoid accidentally deleting merged presentation slides, we could not
            // simply search for Visual Agenda Sections and delete them. Instead, we
            // delete all Visual Agenda Sections without deleting slides first, then
            // delete all transitions slides manually.

            // delete all generated sections without deleting the slides first
            for (var i = section.Count; i >= 1; i--)
            {
                if (section[i - 1] == PptLabsAgendaVisualSectionName)
                {
                    sectionProperties.Delete(i, false);
                }
            }

            // delete all transition slides
            var nameSearchRegex = new Regex("PptLabsVisualAgenda|PPTLabsZoom");
            PowerPointPresentation.Current.RemoveSlide(nameSearchRegex, true);
        }

        private static bool SectionValidation()
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

        private static void SelectOriginalSlide(PowerPointSlide oriSlide, PowerPointSlide defSlide)
        {
            try
            {
                oriSlide.GetNativeSlide().Select();
            }
            catch (COMException)
            {
                defSlide.GetNativeSlide().Select();
            }
        }

        private static void SyncAgendaBeam(PowerPointSlide refSlide, List<PowerPointSlide> slides)
        {
            var refBeamShape = FindBeamShape(refSlide);

            foreach (var slide in slides)
            {
                SyncSingleAgendaBeam(slide, refBeamShape);
            }
        }

        private static void SyncAgendaBullet(List<string> sections, PowerPointSlide refSlide)
        {
            foreach (var section in sections)
            {
                var start = FindSectionStartSlide(section, Type.Bullet);
                var end = FindSectionEndSlide(section, Type.Bullet);

                SyncSingleAgendaGeneral(refSlide, start, Type.Bullet);
                SyncSingleAgendaGeneral(refSlide, end, Type.Bullet);

                SyncSingleAgendaBullet(refSlide, start);
                SyncSingleAgendaBullet(refSlide, end);
            }
        }

        private static void SyncAgendaVisual(List<string> sections, PowerPointSlide refSlide)
        {
            var currentPresentation = PowerPointPresentation.Current;

            // delete all generated transition slides
            foreach (var slide in currentPresentation.Slides.Where(slide => slide.Name.Contains("PPTLabsZoom")))
            {
                slide.Delete();
            }

            var endOfAgenda = currentPresentation.Slides
                                                 .FirstOrDefault(slide => slide.Name.Contains("EndOfAgenda"));
            var sectionCount = sections.Count;

            SyncSingleAgendaGeneral(refSlide, endOfAgenda, Type.Visual);
            SyncSingleAgendaVisual(endOfAgenda, sections, sectionCount);

            for (var i = 0; i < sectionCount; i++)
            {
                var section = sections[i];
                var endAgenda = FindSectionEndSlide(section, Type.Visual);

                SyncSingleAgendaGeneral(refSlide, endAgenda, Type.Visual);
                SyncSingleAgendaVisual(endAgenda, sections, i);
            }
        }

        private static void SyncSingleAgendaBeam(PowerPointSlide slide, Shape refBeamShape)
        {
            var refSubIems = refBeamShape.GroupItems.Cast<Shape>().ToList();
            var refHighlight = FindBeamHighlight(refSubIems);
            var refNormal = FindBeamNormal(refSubIems);
            var refBackground = refSubIems.FirstOrDefault(shape => shape.Name == PptLabsAgendaBeamBackgroundName);
            var beamShape = FindBeamShape(slide);

            if (beamShape == null) return;

            Utils.Graphics.SyncShape(refBeamShape, beamShape, pickupShapeFormat: false);

            var subItems = beamShape.GroupItems.Cast<Shape>().ToList();
            var section = FindSlideSection(slide);
            var widest = 0f;

            foreach (var item in subItems)
            {
                // specially deal with the background shape
                if (item.Name == PptLabsAgendaBeamBackgroundName)
                {
                    Utils.Graphics.SyncShape(refBackground, item, pickupTextContent: false);
                    item.Width = PowerPointPresentation.Current.SlideWidth;
                }
                else
                {
                    // for all items, we have 2 cases:
                    // 1. if the sections have not been changed, we follow the reference for the layout and format;
                    // 2. if the sections have been changed, we follow only the format, but not the layout
                    var itemText = item.TextFrame.TextRange.Text;

                    if (!_agendaOutdated)
                    {
                        var oldItem = refSubIems.FirstOrDefault(shape => shape.TextFrame.TextRange.Text == itemText);

                        // pick up layout first
                        Utils.Graphics.SyncShape(oldItem, item, pickupShapeFormat: false,
                                                 pickupTextContent: false, pickupTextFormat: false);
                    }

                    Utils.Graphics.SyncShape(itemText == section ? refHighlight : refNormal, item,
                                             pickupShapeBasic: false, pickupTextContent: false);

                    widest = Math.Max(widest, item.Width);
                }
            }

            if (_agendaOutdated)
            {
                var lastLeft = 0f;
                var lastTop = 0f;

                var delta = Math.Max(widest, PowerPointPresentation.Current.SlideWidth / (subItems.Count - 1));
                var background = subItems.FirstOrDefault(shape => shape.Name == PptLabsAgendaBeamBackgroundName);

                if (background != null)
                {
                    background.Height = 0;
                }

                foreach (var item in subItems.Where(shape => shape.Name != PptLabsAgendaBeamBackgroundName))
                {
                    // TODO: adjust previous line when lastTop increases
                    AdjustBeamItemHorizontal(ref lastLeft, ref lastTop, ref widest, delta, item, background);
                }
            }
        }

        private static void SyncSingleAgendaBullet(PowerPointSlide refSlide, PowerPointSlide candidate)
        {
            if (refSlide == null || candidate == null)
            {
                return;
            }

            var refContentShape = refSlide.GetShapeWithName(PptLabsAgendaContentShapeName)[0];
            var candidateContentShape = candidate.GetShapeWithName(PptLabsAgendaContentShapeName)[0];

            Utils.Graphics.SyncShape(refContentShape, candidateContentShape, pickupShapeFormat: false, pickupTextContent: false);
        }

        private static void SyncSingleAgendaVisual(PowerPointSlide candidate, List<string> sections, int sectionIndex)
        {
            if (sectionIndex == 0)
            {
                var shape = candidate.GetShapeWithName(PptLabsAgendaVisualItemPrefix + sections[0])[0];
                var endSlideName = string.Format("{0} Start.png", sections[0]);
                shape.Fill.UserPicture(Path.Combine(SlideCapturePath, endSlideName));
            }

            for (var i = 0; i < sectionIndex; i ++)
            {
                var shape = candidate.GetShapeWithName(PptLabsAgendaVisualItemPrefix + sections[i])[0];
                var endSlideName = string.Format("{0} End.png", sections[i]);
                shape.Fill.UserPicture(Path.Combine(SlideCapturePath, endSlideName));
            }
            
            if (sectionIndex < sections.Count)
            {
                GenerateVisualAgendaSlideZoomIn(candidate, sections[sectionIndex]);
            }

            if (sectionIndex > 0)
            {
                GenerateVisualAgendaSlideZoomOut(candidate, sections[sectionIndex - 1], (sectionIndex + 1) * 2);
            }
        }

        private static void SyncSingleAgendaGeneral(PowerPointSlide refSlide, PowerPointSlide candidate, Type type)
        {
            // in this step, we should sync:
            // 1. Layout
            // 2. Design;
            // -3. Transition; -> this is no longer synced because each agenda may have different transition
            // 4. Shapes and their position, text;

            if (refSlide == null || candidate == null ||
                refSlide == candidate)
            {
                return;
            }

            candidate.Layout = refSlide.Layout;
            candidate.Design = refSlide.Design;

            // syncronize extra shapes other than visual items in reference slide
            var extraShapes = refSlide.Shapes.Cast<Shape>()
                                             .Where(shape => !candidate.HasShapeWithSameName(shape.Name) &&
                                                             !shape.Name.Contains(PptLabsAgendaVisualItemPrefix))
                                             .Select(shape => shape.Name)
                                             .ToArray();

            if (extraShapes.Length != 0)
            {
                var refShapes = refSlide.Shapes.Range(extraShapes);
                refShapes.Copy();
                var copiedShapes = candidate.Shapes.Paste();

                Utils.Graphics.SyncShapeRange(refShapes, copiedShapes);
            }

            // sync shape with same name only for bullet agenda, or visual agenda that is still up to date
            if (type == Type.Bullet ||
               (type == Type.Visual && !_agendaOutdated))
            {
                // syncronize shapes position and size, except bullet content
                var sameShapes = refSlide.Shapes.Cast<Shape>()
                                                .Where(shape => shape.Name != PptLabsAgendaContentShapeName &&
                                                                candidate.HasShapeWithSameName(shape.Name));

                foreach (var refShape in sameShapes)
                {
                    var candidateShape = candidate.GetShapeWithName(refShape.Name)[0];

                    Utils.Graphics.SyncShape(refShape, candidateShape);
                }
            }
        }
        # endregion

        # region Event Handler
        private static void UpdateColorScheme(Color highlightColor, Color dimColor, Color defaultColor)
        {
            var sections = PowerPointPresentation.Current.Sections.Skip(1).ToList();
            var type = CurrentType;

            if (type != Type.Bullet) return;

            // take care of the section update
            var refSlide = FindReferenceSlide(type);
            PrepareSync(type, ref refSlide);
            
            // update color settings
            _bulletHighlightColor = highlightColor;
            _bulletDimColor = dimColor;
            _bulletDefaultColor = defaultColor;

            GenerateBulletAgenda(sections);
            SyncAgendaBullet(sections, refSlide);

            refSlide.Delete();
        }
        # endregion
    }
}
