using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Diagnostics;
using System.Windows.Forms;
using PowerPointLabs.Models;
using PowerPointLabs.Views;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

// TODO:  Follow these steps to enable the Ribbon (XML) item:

// 1: Copy the following code block into the ThisAddin, ThisWorkbook, or ThisDocument class.

//  protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
//  {
//      return new Ribbon1();
//  }

// 2. Create callback methods in the "Ribbon Callbacks" region of this class to handle user
//    actions, such as clicking a button. Note: if you have exported this Ribbon from the Ribbon designer,
//    move your code from the event handlers to the callback methods and modify the code to work with the
//    Ribbon extensibility (RibbonX) programming model.

// 3. Assign attributes to the control tags in the Ribbon XML file to identify the appropriate callback methods in your code.  

// For more information, see the Ribbon XML documentation in the Visual Studio Tools for Office Help.


namespace PowerPointLabs
{
    [ComVisible(true)]
    public class Ribbon1 : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI ribbon;
        
        public bool frameAnimationChecked = false;
        public bool backgroundZoomChecked = true;
        public bool multiSlideZoomChecked = false;
        public bool spotlightDelete = true;
        public float defaultDuration = 0.5f;
        
        public bool spotlightEnabled = false;
        public bool inSlideEnabled = false;
        public bool zoomButtonEnabled = false;
        public bool highlightBulletsEnabled = true;
        public bool addAutoMotionEnabled = true;
        public bool reloadAutoMotionEnabled = true;
        public bool reloadSpotlight = true;

        public bool _recorderPaneVisible = false;
        private bool _firstLoadRecorder = true;

        private bool _allSlides;
        private bool _previewCurrentSlide;
        private bool _captionsAllSlides;

        private List<string> _voiceNames;

        private int _voiceSelected = 0;

        public Ribbon1()
        {
        }

        #region IRibbonExtensibility Members

        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText("PowerPointLabs.Ribbon1.xml");
        }

        #endregion

        #region Ribbon Callbacks
        //Create callback methods here. For more information about adding callback methods, select the Ribbon XML item in Solution Explorer and then press F1

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            this.ribbon = ribbonUI;

            SetVoicesFromInstalledOptions();
            SetCoreVoicesToSelections();
        }

        public void RefreshRibbonControl(String controlID)
        {
            try
            {
                ribbon.InvalidateControl(controlID);
            }
            catch (Exception e)
            {
                PowerPointLabsGlobals.LogException(e, "RefreshRibbonControl");
                throw;
            }
        }

        private void SetVoicesFromInstalledOptions()
        {
            var installedVoices = NotesToAudio.GetVoices().ToList();
            _voiceNames = installedVoices;
        }

        public void HighlightBulletsBackgroundButtonClick(Office.IRibbonControl control)
        {
            try
            {
                if (Globals.ThisAddIn.Application.ActiveWindow.Selection.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
                    HighlightBulletsBackground.userSelection = HighlightBulletsBackground.HighlightBackgroundSelection.kShapeSelected;
                else if (Globals.ThisAddIn.Application.ActiveWindow.Selection.Type == PowerPoint.PpSelectionType.ppSelectionText)
                    HighlightBulletsBackground.userSelection = HighlightBulletsBackground.HighlightBackgroundSelection.kTextSelected;
                else
                    HighlightBulletsBackground.userSelection = HighlightBulletsBackground.HighlightBackgroundSelection.kNoneSelected;

                HighlightBulletsBackground.AddHighlightBulletsBackground();
            }
            catch (Exception e)
            {
                PowerPointLabsGlobals.LogException(e, "HighlightBulletsBackgroundButtonClick");
                throw;
            }
        }

        public void HighlightBulletsTextButtonClick(Office.IRibbonControl control)
        {
            try
            {
                if (Globals.ThisAddIn.Application.ActiveWindow.Selection.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
                    HighlightBulletsText.userSelection = HighlightBulletsText.HighlightTextSelection.kShapeSelected;
                else if (Globals.ThisAddIn.Application.ActiveWindow.Selection.Type == PowerPoint.PpSelectionType.ppSelectionText)
                    HighlightBulletsText.userSelection = HighlightBulletsText.HighlightTextSelection.kTextSelected;
                else
                    HighlightBulletsText.userSelection = HighlightBulletsText.HighlightTextSelection.kNoneSelected;

                HighlightBulletsText.AddHighlightBulletsText();
            }
            catch (Exception e)
            {
                PowerPointLabsGlobals.LogException(e, "HighlightBulletsTextButtonClick");
                throw;
            }
        }

        public void AddInSlideAnimationButtonClick(Office.IRibbonControl control)
        {
            try
            {
                AnimateInSlide.isHighlightBullets = false;
                AnimateInSlide.AddAnimationInSlide();
            }
            catch (Exception e)
            {
                PowerPointLabsGlobals.LogException(e, "AddInSlideAnimationButtonClick");
                throw;
            }
        }
        public void ReloadSpotlightButtonClick(Office.IRibbonControl control)
        {
            try
            {
                Spotlight.ReloadSpotlightEffect();
            }
            catch (Exception e)
            {
                PowerPointLabsGlobals.LogException(e, "ReloadSpotlightButtonClick");
                throw;
            }
        }
        public void SpotlightBtnClick(Office.IRibbonControl control)
        {
            try
            {
                Spotlight.AddSpotlightEffect();
            }
            catch (Exception e)
            {
                PowerPointLabsGlobals.LogException(e, "SpotlightBtnClick");
                throw;
            }
        }

        //Button Click Callbacks        
        public void AddAnimationButtonClick(Office.IRibbonControl control)
        {
            try
            {
                AutoAnimate.AddAutoAnimation();
            }
            catch (Exception e)
            {
                PowerPointLabsGlobals.LogException(e, "AddAnimationButtonClick");
                throw;
            }
        }
        public void ReloadButtonClick(Office.IRibbonControl control)
        {
            try
            {
                AutoAnimate.ReloadAutoAnimation();
            }
            catch (Exception e)
            {
                PowerPointLabsGlobals.LogException(e, "ReloadAnimationButtonClick");
                throw;
            }
        }
        public void ZoomBtnClick(Office.IRibbonControl control)
        {
            ZoomToArea.AddZoomToArea();
        }
        public void AboutButtonClick(Office.IRibbonControl control)
        {
            System.Windows.Forms.MessageBox.Show("          PowerPointLabs Plugin Version 1.7.2 [Release date: 3 Apr 2014]\n     Developed at School of Computing, National University of Singapore.\n        For more information, visit our website http://PowerPointLabs.info", "About PowerPointLabs");
        }
        public void HelpButtonClick(Office.IRibbonControl control)
        {
            try
            {
                string myURL = "http://powerpointlabs.info/docs.html";
                System.Diagnostics.Process.Start(myURL);
            }
            catch (Exception e)
            {
                PowerPointLabsGlobals.LogException(e, "HelpButtonClick");
                throw;
            }
        }
        public void FeedbackButtonClick(Office.IRibbonControl control)
        {
            try
            {
                string myURL = "http://powerpointlabs.info/contact.html";
                System.Diagnostics.Process.Start(myURL);
            }
            catch (Exception e)
            {
                PowerPointLabsGlobals.LogException(e, "FeedbackButtonClick");
                throw;
            }
        }
        public void AddZoomInButtonClick(Office.IRibbonControl control)
        {
            try
            {
                AutoZoom.AddDrillDownAnimation();
            }
            catch (Exception e)
            {
                PowerPointLabsGlobals.LogException(e, "AddZoomInButtonClick");
                throw;
            }
        }
        public void AddZoomOutButtonClick(Office.IRibbonControl control)
        {
            try
            {
                AutoZoom.AddStepBackAnimation();
            }
            catch (Exception e)
            {
                PowerPointLabsGlobals.LogException(e, "AddZoomOutButtonClick");
                throw;
            }
        }

        public void CropShapeButtonClick(Office.IRibbonControl control)
        {
            var selection = Globals.ThisAddIn.Application.
                ActiveWindow.Selection;
            CropShape(ref selection);
        }

        public System.Drawing.Bitmap GetAddAnimationImage(Office.IRibbonControl control)
        {
            try
            {
                return new System.Drawing.Bitmap(Properties.Resources.AddAnimation);
            }
            catch (Exception e)
            {
                PowerPointLabsGlobals.LogException(e, "GetAddAnimationImage");
                throw;
            }
        }
        public System.Drawing.Bitmap GetReloadAnimationImage(Office.IRibbonControl control)
        {
            try
            {
                return new System.Drawing.Bitmap(Properties.Resources.ReloadAnimation);
            }
            catch (Exception e)
            {
                PowerPointLabsGlobals.LogException(e, "GetReloadAnimationImage");
                throw;
            }
        }
        public System.Drawing.Bitmap GetSpotlightImage(Office.IRibbonControl control)
        {
            try
            {
                return new System.Drawing.Bitmap(Properties.Resources.Spotlight);
            }
            catch (Exception e)
            {
                PowerPointLabsGlobals.LogException(e, "GetSpotlightImage");
                throw;
            }
        }
        public System.Drawing.Bitmap GetReloadSpotlightImage(Office.IRibbonControl control)
        {
            try
            {
                return new System.Drawing.Bitmap(Properties.Resources.ReloadSpotlight);
            }
            catch (Exception e)
            {
                PowerPointLabsGlobals.LogException(e, "GetReloadSpotlightImage");
                throw;
            }
        }
        public System.Drawing.Bitmap GetHighlightBulletsTextImage(Office.IRibbonControl control)
        {
            try
            {
                return new System.Drawing.Bitmap(Properties.Resources.HighlightText);
            }
            catch (Exception e)
            {
                PowerPointLabsGlobals.LogException(e, "GetHighlightBulletsTextImage");
                throw;
            }
        }
        public System.Drawing.Bitmap GetHighlightBulletsBackgroundImage(Office.IRibbonControl control)
        {
            try
            {
                return new System.Drawing.Bitmap(Properties.Resources.HighlightBackground);
            }
            catch (Exception e)
            {
                PowerPointLabsGlobals.LogException(e, "GetHighlightBulletsBackgroundImage");
                throw;
            }
        }

        public System.Drawing.Bitmap GetHighlightBulletsTextContextImage(Office.IRibbonControl control)
        {
            try
            {
                return new System.Drawing.Bitmap(Properties.Resources.HighlightTextContext);
            }
            catch (Exception e)
            {
                PowerPointLabsGlobals.LogException(e, "GetHighlightBulletsTextContextImage");
                throw;
            }
        }
        public System.Drawing.Bitmap GetHighlightBulletsBackgroundContextImage(Office.IRibbonControl control)
        {
            try
            {
                return new System.Drawing.Bitmap(Properties.Resources.HighlightBackgroundContext);
            }
            catch (Exception e)
            {
                PowerPointLabsGlobals.LogException(e, "GetHighlightBulletsBackgroundContextImage");
                throw;
            }
        }

        public System.Drawing.Bitmap GetZoomInImage(Office.IRibbonControl control)
        {
            try
            {
                return new System.Drawing.Bitmap(Properties.Resources.ZoomIn);
            }
            catch (Exception e)
            {
                PowerPointLabsGlobals.LogException(e, "GetZoomInImage");
                throw;
            }
        }

        public System.Drawing.Bitmap GetZoomOutImage(Office.IRibbonControl control)
        {
            try
            {
                return new System.Drawing.Bitmap(Properties.Resources.ZoomOut);
            }
            catch (Exception e)
            {
                PowerPointLabsGlobals.LogException(e, "GetZoomOutImage");
                throw;
            }
        }
        public System.Drawing.Bitmap GetZoomToAreaImage(Office.IRibbonControl control)
        {
            try
            {
                return new System.Drawing.Bitmap(Properties.Resources.ZoomToArea);
            }
            catch (Exception e)
            {
                PowerPointLabsGlobals.LogException(e, "GetZoomToAreaImage");
                throw;
            }
        }
        public System.Drawing.Bitmap GetZoomToAreaContextImage(Office.IRibbonControl control)
        {
            try
            {
                return new System.Drawing.Bitmap(Properties.Resources.ZoomToAreaContext);
            }
            catch (Exception e)
            {
                PowerPointLabsGlobals.LogException(e, "GetZoomToAreaContextImage");
                throw;
            }
        }
        public System.Drawing.Bitmap GetCropShapeImage(Office.IRibbonControl control)
        {
            try
            {
                return new System.Drawing.Bitmap(Properties.Resources.CutOutShape);
            }
            catch (Exception e)
            {
                PowerPointLabsGlobals.LogException(e, "GetCropShapeImage");
                throw;
            }
        }
        public System.Drawing.Bitmap GetAboutImage(Office.IRibbonControl control)
        {
            try
            {
                return new System.Drawing.Bitmap(Properties.Resources.About);
            }
            catch (Exception e)
            {
                PowerPointLabsGlobals.LogException(e, "GetAboutImage");
                throw;
            }
        }
        public System.Drawing.Bitmap GetHelpImage(Office.IRibbonControl control)
        {
            try
            {
                return new System.Drawing.Bitmap(Properties.Resources.Help);
            }
            catch (Exception e)
            {
                PowerPointLabsGlobals.LogException(e, "GetHelpImage");
                throw;
            }
        }
        public System.Drawing.Bitmap GetFeedbackImage(Office.IRibbonControl control)
        {
            try
            {
                return new System.Drawing.Bitmap(Properties.Resources.Feedback);
            }
            catch (Exception e)
            {
                PowerPointLabsGlobals.LogException(e, "GetFeedbackImage");
                throw;
            }
        }
        public System.Drawing.Bitmap GetAddAudioImage(Office.IRibbonControl control)
        {
            try
            {
                return new System.Drawing.Bitmap(Properties.Resources.AddAudio);
            }
            catch (Exception e)
            {
                PowerPointLabsGlobals.LogException(e, "GetAddAudioImage");
                throw;
            }
        }
        public System.Drawing.Bitmap GetRemoveAudioImage(Office.IRibbonControl control)
        {
            try
            {
                return new System.Drawing.Bitmap(Properties.Resources.RemoveAudio);
            }
            catch (Exception e)
            {
                PowerPointLabsGlobals.LogException(e, "GetRemoveAudioImage");
                throw;
            }
        }
        public System.Drawing.Bitmap GetAddCaptionsImage(Office.IRibbonControl control)
        {
            try
            {
                return new System.Drawing.Bitmap(Properties.Resources.AddCaption);
            }
            catch (Exception e)
            {
                PowerPointLabsGlobals.LogException(e, "GetAddCaptionsImage");
                throw;
            }
        }
        public System.Drawing.Bitmap GetRemoveCaptionsImage(Office.IRibbonControl control)
        {
            try
            {
                return new System.Drawing.Bitmap(Properties.Resources.RemoveCaption);
            }
            catch (Exception e)
            {
                PowerPointLabsGlobals.LogException(e, "GetRemoveCaptionsImage");
                throw;
            }
        }

        public System.Drawing.Bitmap GetAddAudioContextImage(Office.IRibbonControl control)
        {
            try
            {
                return new System.Drawing.Bitmap(Properties.Resources.AddNarrationContext);
            }
            catch (Exception e)
            {
                PowerPointLabsGlobals.LogException(e, "GetAddAudioContextImage");
                throw;
            }
        }
        public System.Drawing.Bitmap GetPreviewNarrationContextImage(Office.IRibbonControl control)
        {
            try
            {
                return new System.Drawing.Bitmap(Properties.Resources.SpeakTextContext);
            }
            catch (Exception e)
            {
                PowerPointLabsGlobals.LogException(e, "GetPreviewNarrationContextImage");
                throw;
            }
        }
        public System.Drawing.Bitmap GetInSlideAnimationImage(Office.IRibbonControl control)
        {
            try
            {
                return new System.Drawing.Bitmap(Properties.Resources.InSlideAnimation);
            }
            catch (Exception e)
            {
                PowerPointLabsGlobals.LogException(e, "GetInSlideAnimationImage");
                throw;
            }
        }
        public System.Drawing.Bitmap GetAddAnimationContextImage(Office.IRibbonControl control)
        {
            try
            {
                return new System.Drawing.Bitmap(Properties.Resources.AddAnimationContext);
            }
            catch (Exception e)
            {
                PowerPointLabsGlobals.LogException(e, "GetAddAnimationContextImage");
                throw;
            }
        }
        public System.Drawing.Bitmap GetReloadAnimationContextImage(Office.IRibbonControl control)
        {
            try
            {
                return new System.Drawing.Bitmap(Properties.Resources.ReloadAnimationContext);
            }
            catch (Exception e)
            {
                PowerPointLabsGlobals.LogException(e, "GetReloadAnimationContextImage");
                throw;
            }
        }
        public System.Drawing.Bitmap GetAddSpotlightContextImage(Office.IRibbonControl control)
        {
            try
            {
                return new System.Drawing.Bitmap(Properties.Resources.AddSpotlightContext);
            }
            catch (Exception e)
            {
                PowerPointLabsGlobals.LogException(e, "GetAddSpotlightContextImage");
                throw;
            }
        }
        public System.Drawing.Bitmap GetEditNameContextImage(Office.IRibbonControl control)
        {
            try
            {
                return new System.Drawing.Bitmap(Properties.Resources.EditNameContext);
            }
            catch (Exception e)
            {
                PowerPointLabsGlobals.LogException(e, "GetEditNameContextImage");
                throw;
            }
        }
        public System.Drawing.Bitmap GetInSlideAnimationContextImage(Office.IRibbonControl control)
        {
            try
            {
                return new System.Drawing.Bitmap(Properties.Resources.InSlideContext);
            }
            catch (Exception e)
            {
                PowerPointLabsGlobals.LogException(e, "GetInSlideAnimationContextImage");
                throw;
            }
        }
        public System.Drawing.Bitmap GetZoomInContextImage(Office.IRibbonControl control)
        {
            try
            {
                return new System.Drawing.Bitmap(Properties.Resources.ZoomInContext);
            }
            catch (Exception e)
            {
                PowerPointLabsGlobals.LogException(e, "GetZoomInContextImage");
                throw;
            }
        }
        public System.Drawing.Bitmap GetZoomOutContextImage(Office.IRibbonControl control)
        {
            try
            {
                return new System.Drawing.Bitmap(Properties.Resources.ZoomOutContext);
            }
            catch (Exception e)
            {
                PowerPointLabsGlobals.LogException(e, "GetZoomOutContextImage");
                throw;
            }
        }
        public void ZoomStyleChanged(Office.IRibbonControl control, bool pressed)
        {
            try
            {
                if (pressed)
                {
                    backgroundZoomChecked = true;
                }
                else
                {
                    backgroundZoomChecked = false;
                }
            }
            catch (Exception e)
            {
                PowerPointLabsGlobals.LogException(e, "ZoomStyleChanged");
                throw;
            }
        }
        public bool ZoomStyleGetPressed(Office.IRibbonControl control)
        {
            try
            {
                return backgroundZoomChecked;
            }
            catch (Exception e)
            {
                PowerPointLabsGlobals.LogException(e, "ZoomStyleGetPressed");
                throw;
            }
        }
        //Control Enabled Callbacks
        public bool OnGetEnabledSpotlight(Office.IRibbonControl control)
        {
            return spotlightEnabled;
        }
        public bool OnGetEnabledReloadSpotlight(Office.IRibbonControl control)
        {
            return reloadSpotlight;
        }
        public bool OnGetEnabledAddAutoMotion(Office.IRibbonControl control)
        {
            return addAutoMotionEnabled;
        }
        public bool OnGetEnabledReloadAutoMotion(Office.IRibbonControl control)
        {
            return reloadAutoMotionEnabled;
        }
        public bool OnGetEnabledAddInSlide(Office.IRibbonControl control)
        {
            return inSlideEnabled;
        }
        public bool OnGetEnabledZoomButton(Office.IRibbonControl control)
        {
            return zoomButtonEnabled;
        }
        public bool OnGetEnabledHighlightBullets(Office.IRibbonControl control)
        {
            return highlightBulletsEnabled;
        }

        //Edit Name Callbacks
        public void NameEditBtnClick(Office.IRibbonControl control)
        {
            try
            {
                PowerPoint.Shape selectedShape = (PowerPoint.Shape)Globals.ThisAddIn.Application.ActiveWindow.Selection.ShapeRange[1];
                Form1 editForm = new Form1(this, selectedShape.Name);
                editForm.ShowDialog();
            }
            catch (Exception e)
            {
                PowerPointLabsGlobals.LogException(e, "NameEditBtnClick");
                throw;
            }
        }
        public void ShapeNameEdited(String newName)
        {
            try
            {
                PowerPoint.Shape selectedShape = (PowerPoint.Shape)Globals.ThisAddIn.Application.ActiveWindow.Selection.ShapeRange[1];
                selectedShape.Name = newName;
            }
            catch (Exception e)
            {
                PowerPointLabsGlobals.LogException(e, "ShapeNameEdited");
                throw;
            }
        }

        public void AutoAnimateDialogButtonPressed(Office.IRibbonControl control)
        {
            try
            {
                AutoAnimateDialogBox dialog = new AutoAnimateDialogBox(defaultDuration, frameAnimationChecked);
                dialog.SettingsHandler += AnimationPropertiesEdited;
                dialog.ShowDialog();
            }
            catch (Exception e)
            {
                PowerPointLabsGlobals.LogException(e, "AutoAnimateDialogButtonPressed");
                throw;
            }
        }

        public void AnimationPropertiesEdited(float newDuration, bool newFrameChecked)
        {
            try
            {
                defaultDuration = newDuration;
                frameAnimationChecked = newFrameChecked;
                AnimateInSlide.defaultDuration = newDuration;
                AnimateInSlide.frameAnimationChecked = newFrameChecked;
                AutoAnimate.defaultDuration = newDuration;
                AutoAnimate.frameAnimationChecked = newFrameChecked;
            }
            catch (Exception e)
            {
                PowerPointLabsGlobals.LogException(e, "AnimationPropertiesEdited");
                throw;
            }
        }

        public void AutoZoomDialogButtonPressed(Office.IRibbonControl control)
        {
            try
            {
                AutoZoomDialogBox dialog = new AutoZoomDialogBox(backgroundZoomChecked, multiSlideZoomChecked);
                dialog.SettingsHandler += ZoomPropertiesEdited;
                dialog.ShowDialog();
            }
            catch (Exception e)
            {
                PowerPointLabsGlobals.LogException(e, "AutoZoomDialogButtonPressed");
                throw;
            }
        }

        public void ZoomPropertiesEdited(bool backgroundChecked, bool multiSlideChecked)
        {
            try
            {
                backgroundZoomChecked = backgroundChecked;
                multiSlideZoomChecked = multiSlideChecked;
                AutoZoom.backgroundZoomChecked = backgroundChecked;
                ZoomToArea.backgroundZoomChecked = backgroundChecked;
                ZoomToArea.multiSlideZoomChecked = multiSlideChecked;
            }
            catch (Exception e)
            {
                PowerPointLabsGlobals.LogException(e, "ZoomPropertiesEdited");
                throw;
            }
        }

        public void SpotlightDialogButtonPressed(Office.IRibbonControl control)
        {
            try
            {
                SpotlightDialogBox dialog = new SpotlightDialogBox(Spotlight.defaultTransparency, Spotlight.defaultSoftEdges);
                dialog.SettingsHandler += SpotlightPropertiesEdited;
                dialog.ShowDialog();
            }
            catch (Exception e)
            {
                PowerPointLabsGlobals.LogException(e, "SpotlightDialogButtonPressed");
                throw;
            }
        }

        public void SpotlightPropertiesEdited(float newTransparency, float newSoftEdge)
        {
            try
            {
                Spotlight.defaultTransparency = newTransparency;
                Spotlight.defaultSoftEdges = newSoftEdge;
            }
            catch (Exception e)
            {
                PowerPointLabsGlobals.LogException(e, "SpotlightPropertiesEdited");
                throw;
            }
        }

        public void HighlightBulletsPropertiesEdited(Color newHighlightColor, Color newDefaultColor, Color newBackgroundColor)
        {
            try
            {
                HighlightBulletsText.highlightColor = newHighlightColor;
                HighlightBulletsText.defaultColor = newDefaultColor;
                HighlightBulletsBackground.backgroundColor = newBackgroundColor;

            }
            catch (Exception e)
            {
                PowerPointLabsGlobals.LogException(e, "HighlightBulletsPropertiesEdited");
                throw;
            }
        }
        public void HighlightBulletsDialogBoxPressed(Office.IRibbonControl control)
        {
            try
            {
                HighlightBulletsDialogBox dialog = new HighlightBulletsDialogBox(HighlightBulletsText.highlightColor, HighlightBulletsText.defaultColor, HighlightBulletsBackground.backgroundColor);
                dialog.SettingsHandler += HighlightBulletsPropertiesEdited;
                dialog.ShowDialog();
            }
            catch (Exception e)
            {
                PowerPointLabsGlobals.LogException(e, "HighlightBulletsDialogBoxPressed");
                throw;
            }
        }

        # region AudioRecord Button Callbacks
        public void AddRecordClick(Office.IRibbonControl control)
        {
            // TODO:
            // Handle exception when user clicks the button without selecting any slides

            // if currently the pane is hidden, show the pane
            if (!_recorderPaneVisible)
            {
                // fire the pane visble change event
                Globals.ThisAddIn.customTaskPane.Visible = true;
                
                // reload the pane
                Globals.ThisAddIn.recorderTaskPane.RecorderPaneReload();
            }
        }
        # endregion

        #region NotesToAudio Button Callbacks
        public void SpeakSelectedTextClick(Office.IRibbonControl control)
        {
            NotesToAudio.SpeakSelectedText();
        }

        public void RemoveAudioClick(Office.IRibbonControl control)
        {
            if (_allSlides)
            {
                NotesToAudio.RemoveAudioFromAllSlides();
            }
            else
            {
                NotesToAudio.RemoveAudioFromCurrentSlide();
            }
        }

        public void AddAudioClick(Office.IRibbonControl control)
        {
            // TODO:
            // Handle exception when user clicks the button without selecting any slides
            var currentSlide = PowerPointPresentation.CurrentSlide;

            if (_allSlides)
            {
                var allAudioFiles = NotesToAudio.EmbedAllSlideNotes();

                // initialize all slides' audio
                Globals.ThisAddIn.recorderTaskPane.InitializeAudioAndScript(allAudioFiles, true);
            }
            else
            {
                var audioFiles = NotesToAudio.EmbedCurrentSlideNotes();

                // initialize the current slide's audio
                Globals.ThisAddIn.recorderTaskPane.InitializeAudioAndScript(currentSlide.ID, audioFiles, true);
            }

            // if current list is visible, update the pane immediately
            if (_recorderPaneVisible)
            {
                Globals.ThisAddIn.recorderTaskPane.UpdateLists(currentSlide.ID);
            }

            PreviewAnimationsIfChecked();
        }

        public void ContextAddAudioClick(Office.IRibbonControl control)
        {
            NotesToAudio.EmbedCurrentSlideNotes();
            PreviewAnimationsIfChecked();
        }
        #endregion

        #region NotesToCaptions Button Callbacks

        public void AddCaptionClick(Office.IRibbonControl control)
        {
            if (_captionsAllSlides)
            {
                NotesToCaptions.EmbedCaptionsOnAllSlides();
            }
            else
            {
                NotesToCaptions.EmbedCaptionsOnCurrentSlide();
            }
        }

        public void RemoveCaptionClick(Office.IRibbonControl control)
        {
            if (_captionsAllSlides)
            {
                NotesToCaptions.RemoveCaptionsFromAllSlides();
            }
            else
            {
                NotesToCaptions.RemoveCaptionsFromCurrentSlide();
            }
        }

        public void ContextReplaceAudioClick(Office.IRibbonControl control)
        {
            NotesToAudio.ReplaceSelectedAudio();
        }

        #endregion

        #region NotesToAudio/Caption Helpers
        public void AutoNarrateDialogButtonPressed(Office.IRibbonControl control)
        {
            try
            {
                var dialog = new AutoNarrateDialogBox(_voiceSelected, _voiceNames, _allSlides,
                    _previewCurrentSlide);
                dialog.SettingsHandler += AutoNarrateSettingsChanged;
                dialog.ShowDialog();
            }
            catch (Exception e)
            {
                PowerPointLabsGlobals.LogException(e, "AutoNarrateDialogButtonPressed");
                throw;
            }
        }

        public void AutoNarrateSettingsChanged(String voiceName, bool allSlides, bool previewCurrentSlide)
        {
            _allSlides = allSlides;
            _previewCurrentSlide = previewCurrentSlide;
            if (!String.IsNullOrWhiteSpace(voiceName))
            {
                NotesToAudio.SetDefaultVoice(voiceName);
                _voiceSelected = _voiceNames.IndexOf(voiceName);
            }
        }

        public void AutoCaptionDialogButtonPressed(Office.IRibbonControl control)
        {
            try
            {
                var dialog = new AutoCaptionDialogBox(_captionsAllSlides);
                dialog.SettingsHandler += AutoCaptionSettingsChanged;
                dialog.ShowDialog();
            }
            catch (Exception e)
            {
                PowerPointLabsGlobals.LogException(e, "AutoCaptionDialogButtonPressed");
                throw;
            }
        }

        public void AutoCaptionSettingsChanged(bool allSlides)
        {
            _captionsAllSlides = allSlides;
        }

        private void PreviewAnimationsIfChecked()
        {
            if (_previewCurrentSlide)
            {
                NotesToAudio.PreviewAnimations();
            }
        }

        #endregion

        private void SetCoreVoicesToSelections()
        {
            string defaultVoice = GetSelectedVoiceOrNull();
            NotesToAudio.SetDefaultVoice(defaultVoice);
        }

        private string GetSelectedVoiceOrNull()
        {
            string selectedVoice = null;
            try
            {
                selectedVoice = _voiceNames.ToArray()[_voiceSelected];
            }
            catch (IndexOutOfRangeException)
            {
                // No voices are installed.
                // (It should be impossible for the index to be out of range otherwise.)
            }
            return selectedVoice;
        }

        #endregion

        #region feature: Fit To Slide | Fit To Width | Fit To Height

        private const int SelectedShapeIndex = 1;
        private const int TopMost = 0;
        private const int LeftMost = 0;

        public void FitToWidthClick(Office.IRibbonControl control)
        {
            DoFitToWidth();
        }

        public void FitToHeightClick(Office.IRibbonControl control)
        {
            DoFitToHeight();
        }

        private void DoFitToHeight()
        {
            var pageSetup = GetPageSetup();
            var selectedShape = GetSelectedShape();
            float shapeSizeRatio = GetSizeRatio(selectedShape.Height, selectedShape.Width);
            float factor = GetRotationFactorForFitToHeight(ref selectedShape);
            //fit to height
            selectedShape.Height = (pageSetup.SlideHeight / factor);
            selectedShape.Width = selectedShape.Height / shapeSizeRatio;
            //move to centre
            selectedShape.Left = (pageSetup.SlideWidth - selectedShape.Width) / 2;
            selectedShape.Top = TopMost;
            AdjustMoveToCentreForFitToHeight(ref selectedShape);
        }

        private void AdjustMoveToCentreForFitToHeight(ref PowerPoint.Shape shape)
        {
            float adjustLength;
            float rotation = SetupRotationValueForAdjustment(ref shape);
            //Pythagorean theorem
            float diagonal = (float)Math.Sqrt(Math.Pow(shape.Width / 2, 2)
                + Math.Pow(shape.Height / 2, 2));
            //Law of cosines
            float oppositeSideLength = (float)Math.Sqrt((Math.Pow(diagonal, 2) * 2
                * (1 - Math.Cos(rotation))));
            float angle1 = (float)Math.Atan(shape.Width / shape.Height);
            float angle2 = (float)((Math.PI - rotation) / 2);
            //case 1:
            if ((shape.Rotation >= 0 && shape.Rotation <= 90)
                || (shape.Rotation > 270 && shape.Rotation <= 360))
            {
                float targetAngle = (float)(Math.PI - angle1 - angle2);
                adjustLength = (float)(oppositeSideLength * Math.Cos(targetAngle));
            }
            //case 2:
            else/* case: 90 < rotation < 270)*/
            {
                float targetAngle = angle1 - angle2;
                adjustLength = (float)(oppositeSideLength * Math.Cos(targetAngle)) - shape.Height;
            }
            shape.Top += adjustLength;
        }

        private float GetRotationFactorForFitToHeight(ref PowerPoint.Shape shape)
        {
            //set up rotation value
            float rotation = SetupRotationValueForRotateFactor(ref shape);
            //calculate factor for Fit to Height
            float shapeSizeRatio = GetSizeRatio(shape.Height, shape.Width);
            float factor = (float)(Math.Cos(rotation) + Math.Sin(rotation) / shapeSizeRatio);
            return factor;
        }

        private void DoFitToWidth()
        {
            var pageSetup = GetPageSetup();
            var selectedShape = GetSelectedShape();
            float shapeSizeRatio = GetSizeRatio(selectedShape.Height, selectedShape.Width);
            //fit to width
            float factor = GetRotationFactorForFitToWidth(ref selectedShape);
            selectedShape.Height = pageSetup.SlideWidth / factor;
            selectedShape.Width = selectedShape.Height / shapeSizeRatio;
            //move to middle
            selectedShape.Top = (pageSetup.SlideHeight - selectedShape.Height) / 2;
            selectedShape.Left = LeftMost;
            //adjustment for rotation
            AdjustMoveToCentreForFitToWidth(ref selectedShape);
        }

        private void AdjustMoveToCentreForFitToWidth(ref PowerPoint.Shape shape)
        {
            float adjustLength;
            float rotation = SetupRotationValueForAdjustment(ref shape);
            //Pythagorean theorem
            float diagonal = (float)Math.Sqrt(Math.Pow(shape.Width / 2, 2)
                + Math.Pow(shape.Height / 2, 2));
            //Law of cosines
            float oppositeSideLength = (float)Math.Sqrt((Math.Pow(diagonal, 2) * 2
                * (1 - Math.Cos(rotation))));
            float angle1 = (float)Math.Atan(shape.Height / shape.Width);
            float angle2 = (float)((Math.PI - rotation) / 2);
            //case 1:
            if ((shape.Rotation >= 0 && shape.Rotation <= 90)
                || (shape.Rotation > 270 && shape.Rotation <= 360))
            {
                float targetAngle = (float)(Math.PI - angle1 - angle2);
                adjustLength = (float)(oppositeSideLength * Math.Cos(targetAngle));
            }
            //case 2:
            else/* case: 90 < rotation < 270)*/
            {
                float targetAngle = angle1 - angle2;
                adjustLength = (float)(oppositeSideLength * Math.Cos(targetAngle)) - shape.Width;
            }
            shape.Left += adjustLength;
        }

        private float SetupRotationValueForAdjustment(ref PowerPoint.Shape shape)
        {
            float rotation = shape.Rotation;
            if (shape.Rotation > 180 && shape.Rotation <= 360)
            {
                rotation = 360 - shape.Rotation;
            }
            return ConvertDegToRad(rotation);
        }

        private float GetRotationFactorForFitToWidth(ref PowerPoint.Shape shape)
        {
            float rotation = SetupRotationValueForRotateFactor(ref shape);
            //calculate factor for Fit to Height
            float shapeSizeRatio = GetSizeRatio(shape.Height, shape.Width);
            float factor = (float)(Math.Sin(rotation) + Math.Cos(rotation) / shapeSizeRatio);
            return factor;
        }

        private float SetupRotationValueForRotateFactor(ref PowerPoint.Shape shape)
        {
            float rotation;
            if (shape.Rotation == 90.0)
            {
                rotation = shape.Rotation;
            }
            else if (shape.Rotation == 270.0)
            {
                rotation = 360 - shape.Rotation;
            }
            else if ((shape.Rotation > 90 && shape.Rotation <= 180)
                     || (shape.Rotation > 270 && shape.Rotation <= 360))
            {
                rotation = (360 - shape.Rotation) % 90;
            }
            else
            {
                rotation = shape.Rotation % 90;
            }
            return ConvertDegToRad(rotation);
        }

        private float ConvertDegToRad(float rotation)
        {
            rotation = (float)((rotation) * Math.PI / 180); return rotation;
        }

        private float GetSizeRatio(float height, float width)
        {
            return height / width;
        }

        private PowerPoint.PageSetup GetPageSetup()
        {
            return Globals.ThisAddIn.Application.ActivePresentation.PageSetup;
        }

        private PowerPoint.Shape GetSelectedShape()
        {
            return Globals.ThisAddIn.Application.
                   ActiveWindow.Selection.ShapeRange[SelectedShapeIndex];
        }

        public System.Drawing.Bitmap GetFitToWidthImage(Office.IRibbonControl control)
        {
            try
            {
                return new System.Drawing.Bitmap(Properties.Resources.FitToWidth);
            }
            catch (Exception e)
            {
                PowerPointLabsGlobals.LogException(e, "GetFitToWidthImage");
                throw;
            }
        }

        public System.Drawing.Bitmap GetFitToHeightImage(Office.IRibbonControl control)
        {
            try
            {
                return new System.Drawing.Bitmap(Properties.Resources.FitToHeight);
            }
            catch (Exception e)
            {
                PowerPointLabsGlobals.LogException(e, "GetFitToHeightImage");
                throw;
            }
        }

        #endregion

        #region feature: Crop to Shape
        public PowerPoint.Shape CropShapeToSlide(ref PowerPoint.Selection selection)
        {
            try
            {
                IsValidSelectionForCropToShape(ref selection);
                var shape = FormShapeForCropToShape(ref selection);
                TakeScreenshotWithoutShape(ref shape);
                return FillInShapeWithScreenshotOfSlide(shape);
                //return shape;
            }
            catch (Exception)
            {
                MessageBox.Show(GetCropToShapeErrorMessage(), "Unable to crop");
                CropToShapeErrorCode = -1;
                return null;
            }
        }

        private void CropShape(ref PowerPoint.Selection selection)
        {
            try
            {
                IsValidSelectionForCropToShape(ref selection);
                var shape = FormShapeForCropToShape(ref selection);
                TakeScreenshotWithoutShape(ref shape);
                FillInShapeWithScreenshot(shape);
            }
            catch (Exception)
            {
                MessageBox.Show(GetCropToShapeErrorMessage(), "Unable to crop");
                CropToShapeErrorCode = -1;
            }
        }

        private int CropToShapeErrorCode = -1;
        private const int ERROR_SELECTION_COUNT_ZERO = 0;
        private const int ERROR_SELECTION_NON_SHAPE = 1;
        private const int ERROR_EXCEED_SLIDE_BOUND = 2;
        private const int ERROR_ROTATION_NON_ZERO = 3;
        const string OfficeVersion2013 = "15.0";
        const string OfficeVersion2010 = "14.0";

        private string GetCropToShapeErrorMessage()
        {
            switch (CropToShapeErrorCode)
            {
                case ERROR_SELECTION_COUNT_ZERO:
                    return "To start 'Crop To Shape', please select at least one shape.";
                case ERROR_SELECTION_NON_SHAPE:
                    return "'Crop To Shape' only supports shape objects.";
                case ERROR_EXCEED_SLIDE_BOUND:
                    return "Please ensure your shape is within the slide.";
                case ERROR_ROTATION_NON_ZERO:
                    return "In the current version, the 'Crop To Shape' feature does not work if the shape is rotated";
                default:
                    return "Undefined error.";
            }
        }

        private PowerPoint.Shape FillInShapeWithScreenshotOfSlide(PowerPoint.Shape shape)
        {
            if (shape.Type != Office.MsoShapeType.msoGroup)
            {
                ProduceFillInBackground(ref shape);
                shape.Fill.UserPicture(GetPathForFillInBackground());
            }
            else
            {
                using (Bitmap slideImage = (Bitmap)Bitmap.FromFile(GetPathToStore()))
                {
                    foreach (var sh in shape.GroupItems)
                    {
                        var shapeGroupItem = sh as PowerPoint.Shape;
                        ProduceFillInBackground(ref shapeGroupItem, slideImage);
                        shapeGroupItem.Fill.UserPicture(GetPathForFillInBackground());
                    }
                }
            }
            shape.Line.Visible = Office.MsoTriState.msoFalse;
            shape.Copy();
            PowerPoint.Shape returnShape = PowerPointLabsGlobals.GetCurrentSlide().Shapes.Paste()[1];
            shape.Delete();
            return returnShape;
        }

        private void FillInShapeWithScreenshot(PowerPoint.Shape shape)
        {
            if (shape.Type != Office.MsoShapeType.msoGroup)
            {
                ProduceFillInBackground(ref shape);
                shape.Fill.UserPicture(GetPathForFillInBackground());
            }
            else
            {
                using (Bitmap slideImage = (Bitmap)Bitmap.FromFile(GetPathToStore()))
                {
                    foreach (var sh in shape.GroupItems)
                    {
                        var shapeGroupItem = sh as PowerPoint.Shape;
                        ProduceFillInBackground(ref shapeGroupItem, slideImage);
                        shapeGroupItem.Fill.UserPicture(GetPathForFillInBackground());
                    }
                }
            }
            shape.Line.Visible = Office.MsoTriState.msoFalse;
            shape.Copy();
            PowerPointLabsGlobals.GetCurrentSlide().Shapes.Paste();
            shape.Delete();
        }

        private void AdjustFillEffect(ref PowerPoint.Shape shape)
        {
            if (shape.Type == Office.MsoShapeType.msoGroup)
            {
                var range = shape.Ungroup();
                foreach (var o in range)
                {
                    var sh = o as PowerPoint.Shape;
                    var tmpFillEffect = sh.Fill;
                    tmpFillEffect.TextureOffsetX = -sh.Left;
                    tmpFillEffect.TextureOffsetY = -sh.Top;
                }
                shape = range.Group();
            }
        }

        private void ProduceFillInBackground(ref PowerPoint.Shape shape)
        {
            using (Bitmap slideImage = (Bitmap)Bitmap.FromFile(GetPathToStore()))
            {
                float horizontalRatio =
                    (float)(GetDesiredExportWidth() / Globals.ThisAddIn.Application.ActivePresentation.PageSetup.SlideWidth);
                float verticalRatio =
                    (float)(GetDesiredExportHeight() / Globals.ThisAddIn.Application.ActivePresentation.PageSetup.SlideHeight);
                Bitmap croppedImage = KiCut(slideImage,
                    shape.Left * horizontalRatio,
                    shape.Top * verticalRatio,
                    shape.Width * horizontalRatio,
                    shape.Height * verticalRatio);
                croppedImage.Save(GetPathForFillInBackground(), ImageFormat.Png);
            }
        }

        private void ProduceFillInBackground(ref PowerPoint.Shape shape, Bitmap slideImage)
        {
            float horizontalRatio =
                (float)(GetDesiredExportWidth() / Globals.ThisAddIn.Application.ActivePresentation.PageSetup.SlideWidth);
            float verticalRatio =
                (float)(GetDesiredExportHeight() / Globals.ThisAddIn.Application.ActivePresentation.PageSetup.SlideHeight);
            Bitmap croppedImage = KiCut(slideImage,
                shape.Left * horizontalRatio,
                shape.Top * verticalRatio,
                shape.Width * horizontalRatio,
                shape.Height * verticalRatio);
            croppedImage.Save(GetPathForFillInBackground(), ImageFormat.Png);
        }

        public static Bitmap KiCut(Bitmap original, float startX, float startY, float width, float height)
        {
            if (original == null)
            {
                return null;
            }
            if (startX >= original.Width || startY >= original.Height)
            {
                return null;
            }
            try
            {
                Bitmap outputImage = new Bitmap((int)width, (int)height, PixelFormat.Format24bppRgb);

                Graphics inputGraphics = Graphics.FromImage(outputImage);
                inputGraphics.DrawImage(original,
                    new Rectangle(0, 0, (int)width, (int)height),
                    new Rectangle((int)startX, (int)startY, (int)width, (int)height),
                    GraphicsUnit.Pixel);
                inputGraphics.Dispose();

                return outputImage;
            }
            catch
            {
                return null;
            }
        }

        private string GetPathForFillInBackground()
        {
            string path = Environment.GetFolderPath(Environment.SpecialFolder.CommonDocuments);
            return path + "\\currentFillInBg.png";
        }

        private string GetPathToStore()
        {
            string path = Environment.GetFolderPath(Environment.SpecialFolder.CommonDocuments);
            return path + "\\currentSlide.png";
        }

        public bool GetVisibilityForCombineShapes(Office.IRibbonControl control)
        {
            return Globals.ThisAddIn.Application.Version == OfficeVersion2010;
        }

        private PowerPoint.Shape ConvertToPicture(ref PowerPoint.Shape shape)
        {
            float rotation = 0;
            try
            {
                rotation = shape.Rotation;
                shape.Rotation = 0;
            }
            catch (Exception)
            {
                //chart cannot be rotated
            }
            shape.Copy();
            float x = shape.Left;
            float y = shape.Top;
            float width = shape.Width;
            float height = shape.Height;
            shape.Delete();
            var pic = PowerPointLabsGlobals.GetCurrentSlide().Shapes.PasteSpecial(PowerPoint.PpPasteDataType.ppPastePNG)[1];
            pic.Left = x + (width - pic.Width) / 2;
            pic.Top = y + (height - pic.Height) / 2;
            pic.Rotation = rotation;
            pic.Select();
            return pic;
        }

        private bool IsFirstOneOverlapWithSecond(PowerPoint.Shape first, PowerPoint.Shape second)
        {
            if (first == second)
                return false;
            if (!(first.Left > second.Left + second.Width ||
                  first.Left + first.Width < second.Left ||
                  first.Top > second.Top + second.Height ||
                  first.Top + first.Height < second.Top))
            {
                return true;
            }
            return false;
        }

        private void RemoveOverlapItems(PowerPoint.Shape shape)
        {
            var shapeRange = PowerPointLabsGlobals.GetCurrentSlide().Shapes;
            for (int i = 1; i < shapeRange.Count; )
            {
                if (IsFirstOneOverlapWithSecond(shapeRange[i], shape))
                    shapeRange[i].Delete();
                else
                    i++;
            }
        }

        private void TakeScreenshotWithoutShape(ref PowerPoint.Shape shape)
        {
            shape.Visible = Office.MsoTriState.msoFalse;
            TakeScreenshot();
            shape.Visible = Office.MsoTriState.msoTrue;
        }

        private void TakeScreenshot()
        {
            PowerPointLabsGlobals.GetCurrentSlide().Export(GetPathToStore(), "PNG",
                (int)GetDesiredExportWidth(), (int)GetDesiredExportHeight());
        }

        private double GetDesiredExportWidth()
        {
            return Globals.ThisAddIn.Application.ActivePresentation.PageSetup.SlideWidth / 72.0 * 96.0;
        }

        private double GetDesiredExportHeight()
        {
            return Globals.ThisAddIn.Application.ActivePresentation.PageSetup.SlideHeight / 72.0 * 96.0;
        }

        private PowerPoint.Shape FormShapeForCropToShape(ref PowerPoint.Selection selection)
        {
            var oldRange = selection.ShapeRange;
            //some shapes in the nameList cannot be used due to 
            //Powerpoint's 'Delete-Undo' bug: when a shape got deleted or cut, then undo,
            //then we can only read its name/width/height/left/top.. for others, it aborts
            //var nameList = ProduceNameList(PowerPointLabsGlobals.GetCurrentSlide().Shapes.Range());
            //'Cut-Paste' is a normal workaround method for the bug mentioned above
            oldRange.Cut();
            oldRange = PowerPointLabsGlobals.GetCurrentSlide().Shapes.Paste();
            var copyRange = MakeCopyWithSamePosition(oldRange);
            var ungroupedCopyRange = UngroupAllShapes(copyRange);
            PowerPoint.Shape mergedShape = ungroupedCopyRange[1];
            //if (Globals.ThisAddIn.Application.Version == OfficeVersion2013)
            //{
            //    mergedShape = MergeAllShapes(ref ungroupedCopyRange, nameList);
            //    RemoveRotation(ref mergedShape, nameList);
            //}
            //else 
            //if (Globals.ThisAddIn.Application.Version == OfficeVersion2010)
            //{
            if (ungroupedCopyRange.Count > 1)
            {
                mergedShape = ungroupedCopyRange.Group();
            }
            //}
            if (IsWithinSlide(mergedShape))
            {
                oldRange.Delete();
            }
            else
            {
                mergedShape.Delete();
                ThrowExceptionFromCropToShape(ERROR_EXCEED_SLIDE_BOUND);
            }
            return mergedShape;
        }

        private PowerPoint.ShapeRange MakeCopyWithSamePosition(PowerPoint.ShapeRange oldRange)
        {
            //setup shapes' names in oldRange, so that shapes' names in copyRange is the same
            ModifyNames(oldRange);
            oldRange.Copy();
            var copyRange = PowerPointLabsGlobals.GetCurrentSlide().Shapes.Paste();
            //adjust position
            Dictionary<string, Tuple<float, float>> dict = new Dictionary<string, Tuple<float, float>>();
            foreach (var sh in oldRange)
            {
                var shape = sh as PowerPoint.Shape;
                dict.Add(shape.Name, new Tuple<float, float>(shape.Left, shape.Top));
            }
            foreach (var sh in copyRange)
            {
                var shape = sh as PowerPoint.Shape;
                shape.Left = dict[shape.Name].Item1;
                shape.Top = dict[shape.Name].Item2;
            }
            //append something in case we want to differentiate them
            ModifyNames(copyRange, "_Copy");
            return copyRange;
        }

        private List<string> ProduceNameList(PowerPoint.ShapeRange range)
        {
            List<string> output = new List<string>();
            foreach (var sh in range)
            {
                output.Add((sh as PowerPoint.Shape).Name);
            }
            return output;
        }

        private void ModifyNames(PowerPoint.ShapeRange range, string appendString = "")
        {
            if (appendString != "")
            {
                foreach (var sh in range)
                {
                    (sh as PowerPoint.Shape).Name += appendString;
                }
            }
            else
            {
                foreach (var sh in range)
                {
                    (sh as PowerPoint.Shape).Name = Guid.NewGuid().ToString();
                }
            }
        }

        private bool IsWithinSlide(PowerPoint.Shape shape)
        {
            //-1 and +1 for better user experience
            bool cond1 = shape.Left >= -1;
            bool cond2 = shape.Top >= -1;
            bool cond3 = shape.Left + shape.Width <= Globals.ThisAddIn.Application.ActivePresentation.PageSetup.SlideWidth + 1;
            bool cond4 = shape.Top + shape.Height <= Globals.ThisAddIn.Application.ActivePresentation.PageSetup.SlideHeight + 1;
            return cond1 && cond2 && cond3 && cond4;
        }

        private PowerPoint.ShapeRange UngroupAllShapes(PowerPoint.ShapeRange range)
        {
            List<string> ungroupedShapes = new List<string>();
            Queue<PowerPoint.Shape> selectedShapes = new Queue<PowerPoint.Shape>();

            foreach (var item in range)
            {
                selectedShapes.Enqueue(item as PowerPoint.Shape);
            }
            while (selectedShapes.Count != 0)
            {
                var shape = selectedShapes.Dequeue();
                if (shape.Type == Office.MsoShapeType.msoGroup)
                {
                    var subRange = shape.Ungroup();
                    foreach (var item in subRange)
                    {
                        selectedShapes.Enqueue(item as PowerPoint.Shape);
                    }
                }
                else if (/*Globals.ThisAddIn.Application.Version == OfficeVersion2010
                    && */shape.Rotation != 0.0)
                {
                    //remove copies before throwing exception
                    shape.Delete();
                    if (ungroupedShapes.Count > 0)
                    {
                        PowerPointLabsGlobals.GetCurrentSlide().Shapes.Range(ungroupedShapes.ToArray()).Delete();
                    }
                    while (selectedShapes.Count != 0)
                    {
                        selectedShapes.Dequeue().Delete();
                    }
                    ThrowExceptionFromCropToShape(ERROR_ROTATION_NON_ZERO);
                }
                else if (!IsShape(shape))
                {
                    //remove copies before throwing exception
                    shape.Delete();
                    if (ungroupedShapes.Count > 0)
                    {
                        PowerPointLabsGlobals.GetCurrentSlide().Shapes.Range(ungroupedShapes.ToArray()).Delete();
                    }
                    while (selectedShapes.Count != 0)
                    {
                        selectedShapes.Dequeue().Delete();
                    }
                    ThrowExceptionFromCropToShape(ERROR_SELECTION_NON_SHAPE);
                }
                else
                {
                    shape.Name = Guid.NewGuid().ToString();
                    ungroupedShapes.Add(shape.Name);
                }
            }
            return PowerPointLabsGlobals.GetCurrentSlide().Shapes.Range(ungroupedShapes.ToArray());
        }

        //private PowerPoint.Shape MergeAllShapes(ref PowerPoint.ShapeRange range, List<string> nameList)
        //{
        //    PowerPoint.Shape shape = range[1];
        //    if (range.Count > 1)
        //    {
        //        int KaisBirthday = 0x900209;
        //        range[1].Line.ForeColor.RGB = KaisBirthday;
        //        //merged shape will inherit the foreColor of range[1]'s line,
        //        //which is Xie Kai's birthday :p
        //        //a better choice can be: random number or timestamp
        //        range.MergeShapes(Office.MsoMergeCmd.msoMergeUnion, range[1]);
        //        //find the merged shape
        //        var newRange = PowerPointLabsGlobals.GetCurrentSlide().Shapes.Range();
        //        foreach (var sh in newRange)
        //        {
        //            var mergedShape = sh as PowerPoint.Shape;
        //            if (!nameList.Contains(mergedShape.Name))
        //            {
        //                if(KaisBirthday == mergedShape.Line.ForeColor.RGB)
        //                {
        //                    shape = mergedShape;
        //                }
        //            }
        //        }
        //    }
        //    return shape;
        //}

        //private void RemoveRotation(ref PowerPoint.Shape shape, List<string> nameList)
        //{
        //    var helperShape = PowerPointLabsGlobals.GetCurrentSlide().Shapes.AddShape(Office.MsoAutoShapeType.msoShapeRectangle, -100, -100, 1, 1);
        //    int KaisBirthday = 0x900209;
        //    helperShape.Line.ForeColor.RGB = KaisBirthday;
        //    //merged shape will inherit the foreColor of range[1]'s line,
        //    //which is Xie Kai's birthday :p
        //    //a better choice can be: random number or timestamp
        //    var range = PowerPointLabsGlobals.GetCurrentSlide().
        //        Shapes.Range(new List<string> { helperShape.Name, shape.Name }.ToArray());
        //    //Separate the shapes and make rotation back to zero
        //    range.MergeShapes(Office.MsoMergeCmd.msoMergeFragment, helperShape);
        //    //find those resulted shapes
        //    var newRange = PowerPointLabsGlobals.GetCurrentSlide().Shapes.Range();
        //    List<string> list = new List<string>();
        //    foreach (var sh in newRange)
        //    {
        //        var mergedShape = sh as PowerPoint.Shape;
        //        if (mergedShape.Left == -100
        //            && mergedShape.Top == -100
        //            && mergedShape.Width == 1
        //            && mergedShape.Height == 1)
        //        {
        //            helperShape = mergedShape;
        //        }
        //        else if (!nameList.Contains(mergedShape.Name))
        //        {
        //            if(KaisBirthday == mergedShape.Line.ForeColor.RGB)
        //            {
        //                list.Add(mergedShape.Name);
        //            }
        //        }
        //    }
        //    helperShape.Delete();
        //    range = PowerPointLabsGlobals.GetCurrentSlide().Shapes.Range(list.ToArray());
        //    if (list.Count > 1)
        //    {
        //        range.MergeShapes(Office.MsoMergeCmd.msoMergeUnion);
        //        //find out the merged shape
        //        newRange = PowerPointLabsGlobals.GetCurrentSlide().Shapes.Range();
        //        foreach (var sh in newRange)
        //        {
        //            var mergedShape = sh as PowerPoint.Shape;
        //            if (!nameList.Contains(mergedShape.Name))
        //            {
        //                if (KaisBirthday == mergedShape.Line.ForeColor.RGB)
        //                {
        //                    shape = mergedShape;
        //                }
        //            }
        //        }
        //    }
        //    else
        //    {
        //        shape = range[1];
        //    }
        //}

        private void ThrowExceptionFromCropToShape(int typeOfError)
        {
            CropToShapeErrorCode = typeOfError;
            throw new Exception("Error: " + typeOfError.ToString());
        }

        private void IsValidSelectionForCropToShape(ref PowerPoint.Selection selection)
        {
            if (selection.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
            {
                if (selection.ShapeRange.Count < 1)
                {
                    ThrowExceptionFromCropToShape(ERROR_SELECTION_COUNT_ZERO);
                }
                if (!IsShape(selection))
                {
                    ThrowExceptionFromCropToShape(ERROR_SELECTION_NON_SHAPE);
                }
            }
            else
            {
                ThrowExceptionFromCropToShape(ERROR_SELECTION_COUNT_ZERO);
            }
        }

        private bool IsShape(PowerPoint.Selection sel)
        {
            var shapeRange = sel.ShapeRange;
            foreach (var o in shapeRange)
            {
                var shape = o as PowerPoint.Shape;
                if (!IsShape(shape))
                    return false;
            }
            return true;
        }

        private bool IsShape(PowerPoint.Shape shape)
        {
            if (shape.Type != Office.MsoShapeType.msoAutoShape
                    && shape.Type != Office.MsoShapeType.msoFreeform
                    && shape.Type != Office.MsoShapeType.msoGroup)
            {
                return false;
            }
            return true;
        }

        public System.Drawing.Bitmap GetCutOutShapeMenuImage(Office.IRibbonControl control)
        {
            try
            {
                return new System.Drawing.Bitmap(Properties.Resources.CutOutShapeMenu);
            }
            catch (Exception e)
            {
                PowerPointLabsGlobals.LogException(e, "GetCutOutShapeMenuImage");
                throw;
            }
        }

        #endregion

        #region feature: Convert to Picture

        public void ConvertToPictureButtonClick(Office.IRibbonControl control)
        {
            var selection = Globals.ThisAddIn.Application.
                ActiveWindow.Selection;
            if (selection.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
            {
                var shape = selection.ShapeRange[1];
                if (selection.ShapeRange.Count > 1)
                {
                    shape = selection.ShapeRange.Group();
                }
                shape.Cut();
                shape = PowerPointLabsGlobals.GetCurrentSlide().Shapes.Paste()[1];
                ConvertToPicture(ref shape);
            }
            else
            {
                MessageBox.Show("Convert to Picture only supports Shapes and Charts.", "Unable to Convert to Picture");
            }
        }

        public System.Drawing.Bitmap GetConvertToPicMenuImage(Office.IRibbonControl control)
        {
            try
            {
                return new System.Drawing.Bitmap(Properties.Resources.ConvertToPicture);
            }
            catch (Exception e)
            {
                PowerPointLabsGlobals.LogException(e, "GetConvertToPicMenuImage");
                throw;
            }
        }

        #endregion


        private static string GetResourceText(string resourceName)
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            string[] resourceNames = asm.GetManifestResourceNames();
            for (int i = 0; i < resourceNames.Length; ++i)
            {
                if (string.Compare(resourceName, resourceNames[i], StringComparison.OrdinalIgnoreCase) == 0)
                {
                    using (StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(resourceNames[i])))
                    {
                        if (resourceReader != null)
                        {
                            return resourceReader.ReadToEnd();
                        }
                    }
                }
            }
            return null;
        }
    }
}
