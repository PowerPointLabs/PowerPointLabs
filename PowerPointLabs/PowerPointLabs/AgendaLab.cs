using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
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
        private const string PptLabsAgendaSectionName = "PptLabsAgendaSection";
        private const string PptLabsAgendaBeamBackgroundName = "PptLabsAgendaBeamBackground";
        private const string PptLabsAgendaBeamShapeName = "PptLabsAgendaBeamShape";

        private const float VisualAgendaItemMargin = 0.05f;

        private static LoadingDialog _loadDialog = new LoadingDialog();

        private static readonly Regex AgendaSlideSearchPattern = new Regex(PptLabsAgendaSlideTypeSearchPattern);

        private static readonly string SlideCapturePath = Path.Combine(Path.GetTempPath(), "PowerPointLabs Temp");

        private static string _agendaText;

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

        public static void GenerateAgenda(Type type, bool showLoadingDialog = true, bool confirmDelete = true,
                                          List<int> beamId = null)
        {
            // agenda exists in current presentation
            if (confirmDelete && CurrentType != Type.None)
            {
                var confirm = MessageBox.Show(TextCollection.AgendaLabAgendaExistError,
                                              TextCollection.AgendaLabAgendaExistErrorCaption,
                                              MessageBoxButtons.OKCancel);

                if (confirm == DialogResult.OK)
                {
                    RemoveAgenda();
                }
                else
                {
                    return;
                }
            }

            // validate section information
            var sections = PowerPointPresentation.Current.Sections;

            if (sections.Count == 0)
            {
                MessageBox.Show(TextCollection.AgendaLabNoSectionError);
                return;
            }

            if (sections.Count == 1)
            {
                MessageBox.Show(TextCollection.AgendaLabSingleSectionError);
                return;
            }

            sections = sections.Skip(1).ToList();

            if (showLoadingDialog)
            {
                _loadDialog = new LoadingDialog(TextCollection.AgendaLabLoadingDialogTitle,
                                                TextCollection.AgendaLabLoadingDialogContent);
                _loadDialog.Show();
                _loadDialog.Refresh();
            }

            switch (type)
            {
                case Type.Beam:
                    GenerateBeamAgenda(sections, beamId);
                    break;
                case Type.Bullet:
                    GenerateBulletAgenda(sections);
                    break;
                case Type.Visual:
                    GenerateVisualAgenda(sections);
                    break;
            }

            if (showLoadingDialog)
            {
                _loadDialog.Dispose();
            }
        }

        public static void RemoveAgenda(bool all = false)
        {
            var type = CurrentType;

            if (type == Type.None)
            {
                MessageBox.Show(TextCollection.AgendaLabNoAgendaError);
                return;
            }

            switch (type)
            {
                case Type.Beam:
                    RemoveBeamAgenda(all);
                    break;
                case Type.Bullet:
                    RemoveBulletAgenda();
                    break;
                case Type.Visual:
                    RemoveVisualAgenda();
                    break;
            }
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
            _loadDialog = new LoadingDialog("Synchronizing...", "Agenda is getting synchronized, please wait...");
            _loadDialog.Show();
            _loadDialog.Refresh();

            var type = CurrentType;

            if (type == Type.None)
            {
                MessageBox.Show(TextCollection.AgendaLabNoAgendaError);
                return;
            }

            // find the agenda for the first section as reference
            var currentPresentation = PowerPointPresentation.Current;
            var sections = currentPresentation.Sections.Where(section =>
                                                              section != PptLabsAgendaSectionName).Skip(1).ToList();

            var refSlide = FindReferenceSlide(type, sections[0]);

            PrepareSync(type, ref refSlide);

            switch (type)
            {
                case Type.Beam:
                    SyncAgendaBeam(refSlide);
                    break;
                case Type.Bullet:
                    SyncAgendaBullet(sections, refSlide);
                    break;
                case Type.Visual:
                    SyncAgendaVisual(sections, refSlide);
                    break;
            }

            refSlide.Delete();
            _loadDialog.Dispose();
        }

        public static void UpdateBeamAgendaStyle(Direction direction)
        {
            BeamDirection = direction;

            if (CurrentType != Type.Beam) return;

            RemoveAgenda();
            GenerateAgenda(Type.Beam);
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
                currentTextBox.TextFrame.TextRange.Font.Bold = MsoTriState.msoTrue;
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
            var sectionProperties = currentPresentation.SectionProperties;

            for (var i = 0; i <= sections.Count; i++)
            {
                var sectionName = i < sections.Count ? sections[i] : sections[i - 1];
                var prevSectionName = i == 0 ? string.Empty : sections[i - 1];

                // add a new section before the first section and rename to PptLabsAgendaSectionName
                var index = i < sections.Count ? FindSectionStart(sectionName) : PowerPointPresentation.Current.SlideCount;
                var nativeSlide = i == 0 ? currentPresentation.Slides.Add(index, PpSlideLayout.ppLayoutTitleOnly) :
                                           currentPresentation.Slides.Paste(index)[1];
                var slide = PowerPointSlide.FromSlideFactory(nativeSlide);

                slide.Name = string.Format(PptLabsAgendaSlideNameFormat, Type.Visual,
                                           string.Empty, i < sections.Count ? sectionName : "EndOfAgenda");
                var newSectionIndex = sectionProperties.AddBeforeSlide(index, PptLabsAgendaSectionName);

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
            if (lastLeft >= PowerPointPresentation.Current.SlideWidth)
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

        private static string CreateInDocHyperLink(PowerPointSlide slide)
        {
            return slide.ID + "," + slide.Index + "," + slide.Name;
        }

        private static Shape FindBeamShape(PowerPointSlide slide)
        {
            var grouped = slide.GetShapesWithTypeAndRule(MsoShapeType.msoGroup, new Regex(".+"));
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

        private static PowerPointSlide FindReferenceSlide(Type type, string firstSection)
        {
            if (type == Type.Beam)
            {
                var slides = PowerPointPresentation.Current.Slides;

                return slides.FirstOrDefault(slide => slide.GetShapeWithName(PptLabsAgendaBeamShapeName).Count != 0);
            }

            return FindSectionStartSlide(firstSection, type);
        }

        private static int FindSectionEnd(string section)
        {
            var sectionIndex = FindSectionIndex(section);

            return FindSectionEnd(sectionIndex);
        }

        private static int FindSectionEnd(int sectionIndex)
        {
            // here the sectionIndex is 1-based!
            var sectionProperties = PowerPointPresentation.Current.Presentation.SectionProperties;

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
            var sectionProperties = PowerPointPresentation.Current.Presentation.SectionProperties;

            return sectionProperties.FirstSlide(sectionIndex);
        }

        private static PowerPointSlide FindSectionStartSlide(string section, Type type)
        {
            var curPresentation = PowerPointPresentation.Current;
            var slides = curPresentation.Slides;

            if (type == Type.Beam)
            {
                var sectionProperties = curPresentation.Presentation.SectionProperties;
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

        private static void GenerateBeamAgenda(List<string> sections, List<int> beamId = null)
        {
            var firstSectionIndex = FindSectionIndex(sections[0]);
            var slides = PowerPointPresentation.Current.Slides;
            List<PowerPointSlide> selectedSlides;

            if (beamId == null)
            {
                selectedSlides = PowerPointCurrentPresentationInfo.SelectedSlides
                                                                  .Where(slide => slide.Index >= firstSectionIndex)
                                                                  .ToList();
            } else
            {
                selectedSlides = beamId.Select(id => slides.FirstOrDefault(slide => slide.ID == id)).ToList();
            }

            if (selectedSlides.Count < 1) return;

            PrepareBeamAgendaShapes(sections, selectedSlides[0]);

            foreach (var slide in selectedSlides)
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
            var slideShape = slide.GetShapeWithName(sectionName)[0];

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
            var slideShape = slide.GetShapeWithName(sectionName)[0];

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
            var slides = PowerPointPresentation.Current.Presentation.Slides;
            var refSection = FindSlideSection(refSlide);
            // specially handle the beam type removing since we need to update all old slides, but we need
            // to know who are the old slides. Here we record down their id for later reference.
            var beamSlideId =
                PowerPointPresentation.Current.Slides.Where(slide => FindBeamShape(slide) != null).Select(
                    slide => slide.ID).ToList();

            // pick up color setting before we remove the agenda when the type is bullet
            if (type == Type.Bullet)
            {
                PickupColorSettings();
            }
            
            refSlide.GetNativeSlide().Copy();
            var tempRefSlide = PowerPointSlide.FromSlideFactory(slides.Paste(1)[1]);
            tempRefSlide.Design = refSlide.Design;

            RemoveAgenda(true);
            refSlide = tempRefSlide;

            if (type == Type.Beam)
            {
                refSlide.Name = refSection;
                
                // regenerate slides, do not show loading dialog, do not confirm deletion, but generate only
                // for sepcific id
                GenerateAgenda(type, false, false, beamSlideId);
            } else
            {
                // regenerate slides, do not show loading dialog, do not confirm deletion
                GenerateAgenda(type, false, false);
            }
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

                shape.Name = sections[i];
                shape.Line.Visible = MsoTriState.msoFalse;

                var slideCaptureName = string.Format("{0} Start.png", shape.Name);
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

        private static void RemoveBeamAgenda(bool all = false)
        {
            var selectedSlides = all ? PowerPointPresentation.Current.Slides : 
                                       PowerPointCurrentPresentationInfo.SelectedSlides;

            foreach (var selectedSlide in selectedSlides)
            {
                var beamShape = FindBeamShape(selectedSlide);

                if (beamShape != null)
                {
                    beamShape.Delete();
                }
            }
        }

        private static void RemoveBulletAgenda()
        {
            var slides = PowerPointPresentation.Current.Slides;

            foreach (var slide in slides.Where(slide => AgendaSlideSearchPattern.IsMatch(slide.Name)))
            {
                slide.Delete();
            }
        }

        private static void RemoveVisualAgenda()
        {
            var sectionProperties = PowerPointPresentation.Current.Presentation.SectionProperties;
            var section = PowerPointPresentation.Current.Sections;

            for (var i = section.Count; i >= 1; i--)
            {
                if (section[i - 1] == PptLabsAgendaSectionName)
                {
                    sectionProperties.Delete(i, true);
                }
            }
        }

        private static void SyncAgendaBeam(PowerPointSlide refSlide)
        {
            // for refslide, we need to record which section the refslide is in so that
            // we could distinguish highlight item from normal item. Also, we need to 
            // rename the shape here since copy-paste slide will clear slide's name

            var slides = PowerPointPresentation.Current.Slides.Skip(refSlide.Index);
            var refBeamShape = FindBeamShape(refSlide);
            var refSubIems = refBeamShape.GroupItems.Cast<Shape>().ToList();
            var refHighlight = refSubIems.FirstOrDefault(shape => shape.Name == refSlide.Name);
            var refNormal = refSubIems.FirstOrDefault(shape => shape.Name != refSlide.Name &&
                                                               shape.Name != PptLabsAgendaBeamBackgroundName);

            foreach (var slide in slides)
            {
                var beamShape = FindBeamShape(slide);

                if (beamShape == null) continue;

                Utils.Graphics.SyncShape(refBeamShape, beamShape, pickupShapeFormat: false);

                var subItems = beamShape.GroupItems.Cast<Shape>().ToList();
                var section = FindSlideSection(slide);

                foreach (var item in subItems)
                {
                    var correspond = refSubIems.FirstOrDefault(shape => shape.Name == item.Name);

                    Utils.Graphics.SyncShape(correspond, item, pickupTextContent: false);

                    if (item.Name == section)
                    {
                        // sync highlight item
                        Utils.Graphics.SyncShape(refHighlight, item, pickupTextContent: false, pickupShapeBasic: false);
                    }
                    else
                        if (item.Name == refSlide.Name)
                        {
                            // remove highlight item format and change to normal
                            Utils.Graphics.SyncShape(refNormal, item, pickupTextContent: false, pickupShapeBasic: false);
                        }
                }
            }
        }

        private static void SyncAgendaBullet(List<string> sections, PowerPointSlide refSlide)
        {
            foreach (var section in sections)
            {
                var start = FindSectionStartSlide(section, Type.Bullet);
                var end = FindSectionEndSlide(section, Type.Bullet);

                SyncSingleAgendaGeneral(refSlide, start);
                SyncSingleAgendaGeneral(refSlide, end);

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

            SyncSingleAgendaGeneral(refSlide, endOfAgenda);
            SyncSingleAgendaVisual(endOfAgenda, sections, sectionCount);

            for (var i = 0; i < sectionCount; i++)
            {
                var section = sections[i];
                var endAgenda = FindSectionEndSlide(section, Type.Visual);

                SyncSingleAgendaGeneral(refSlide, endAgenda);
                SyncSingleAgendaVisual(endAgenda, sections, i);
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
            if (sectionIndex < sections.Count)
            {
                GenerateVisualAgendaSlideZoomIn(candidate, sections[sectionIndex]);
            }

            if (sectionIndex > 0)
            {
                GenerateVisualAgendaSlideZoomOut(candidate, sections[sectionIndex - 1], (sectionIndex + 1) * 2);
            }
        }

        private static void SyncSingleAgendaGeneral(PowerPointSlide refSlide, PowerPointSlide candidate)
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

            // syncronize extra shapes in reference slide
            var extraShapes = refSlide.Shapes.Cast<Shape>()
                                             .Where(shape => !candidate.HasShapeWithSameName(shape.Name))
                                             .Select(shape => shape.Name)
                                             .ToArray();

            if (extraShapes.Length != 0)
            {
                var refShapes = refSlide.Shapes.Range(extraShapes);
                refShapes.Copy();
                var copiedShapes = candidate.Shapes.Paste();

                Utils.Graphics.SyncShapeRange(refShapes, copiedShapes);
            }

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

        private static void UpdateColorScheme(PowerPointSlide slide, string section, bool isEnd)
        {
            var contentPlaceHolder = slide.GetShapeWithName(PptLabsAgendaContentShapeName)[0];
            var textRange = contentPlaceHolder.TextFrame.TextRange;
            var focusIndex = FindSectionIndex(section) - 1;
            var focusColor = isEnd ? _bulletDimColor : _bulletHighlightColor;

            RecolorTextRange(textRange, focusIndex, focusColor);
        }
        # endregion

        # region Event Handler
        private static void UpdateColorScheme(Color highlightColor, Color dimColor, Color defaultColor)
        {
            _bulletHighlightColor = highlightColor;
            _bulletDimColor = dimColor;
            _bulletDefaultColor = defaultColor;

            var sections = PowerPointPresentation.Current.Sections;
            var type = CurrentType;

            if (type != Type.Bullet) return;

            // skip the default section
            foreach (var section in sections.Skip(1))
            {
                var startAgenda = FindSectionStartSlide(section, type);
                var endAgenda = FindSectionEndSlide(section, type);

                UpdateColorScheme(startAgenda, section, false);
                UpdateColorScheme(endAgenda, section, true);
            }
        }
        # endregion
    }
}
