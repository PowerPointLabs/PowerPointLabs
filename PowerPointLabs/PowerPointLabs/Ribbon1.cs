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
        public bool removeCaptionsEnabled = true;
        
        public bool _recorderPaneVisible = false;
        private bool _firstLoadRecorder = true;

        private bool _allSlides;
        private bool _previewCurrentSlide;
        
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
            MessageBox.Show("          PowerPointLabs Plugin Version 1.7.9 [Release date: 12 Jun 2014]\n     Developed at School of Computing, National University of Singapore.\n        For more information, visit our website http://PowerPointLabs.info", "About PowerPointLabs");
        }
        public void HelpButtonClick(Office.IRibbonControl control)
        {
            try
            {
                string myURL = "http://powerpointlabs.info/docs.html";
                Process.Start(myURL);
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
        public bool OnGetEnabledRemoveCaptions(Office.IRibbonControl control)
        {
            return removeCaptionsEnabled;
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
        public void RecManagementClick(Office.IRibbonControl control)
        {
            var recorder = Globals.ThisAddIn.ActivateCustomTaskPane.Control as RecorderTaskPane;
            // TODO:
            // Handle exception when user clicks the button without selecting any slides

            // if currently the pane is hidden, show the pane
            if (!Globals.ThisAddIn.ActivateCustomTaskPane.Visible)
            {
                // fire the pane visble change event
                Globals.ThisAddIn.ActivateCustomTaskPane.Visible = true;
                
                // reload the pane
                recorder.RecorderPaneReload();
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
            var currentSlide = PowerPointPresentation.CurrentSlide;
            var recorderPane = Globals.ThisAddIn.ActivateCustomTaskPane.Control as RecorderTaskPane;
            
            if (_allSlides)
            {
                NotesToAudio.RemoveAudioFromAllSlides();
                recorderPane.ClearRecordDataList();
            }
            else
            {
                try
                {
                    NotesToAudio.RemoveAudioFromCurrentSlide();
                }
                catch (Exception e)
                {
                    MessageBox.Show(e.Message);
                    throw;
                }
                recorderPane.ClearRecordDataList(currentSlide.ID);
            }

            // if current list is visible, update the pane immediately
            if (Globals.ThisAddIn.ActivateCustomTaskPane.Visible)
            {
                recorderPane.UpdateLists(currentSlide.ID);
            }
        }

        public void AddAudioClick(Office.IRibbonControl control)
        {
            var currentSlide = PowerPointPresentation.CurrentSlide;
            var recorderPane = Globals.ThisAddIn.ActivateCustomTaskPane.Control as RecorderTaskPane;

            if (_allSlides)
            {
                var allAudioFiles = NotesToAudio.EmbedAllSlideNotes();

                // initialize all slides' audio
                recorderPane.InitializeAudioAndScript(allAudioFiles, true);
            }
            else
            {
                var audioFiles = NotesToAudio.EmbedCurrentSlideNotes();

                // initialize the current slide's audio
                recorderPane.InitializeAudioAndScript(currentSlide, audioFiles, true);
            }

            // if current list is visible, update the pane immediately
            if (Globals.ThisAddIn.ActivateCustomTaskPane.Visible)
            {
                recorderPane.UpdateLists(currentSlide.ID);
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
            foreach (PowerPointSlide slide in PowerPointPresentation.SelectedSlides)
            {
                if (slide.NotesPageText.Trim() != "")
                {
                    removeCaptionsEnabled = true;
                    break;
                }
            }
            NotesToCaptions.EmbedCaptionsOnSelectedSlides();
            RefreshRibbonControl("removeCaptions");
        }

        public void RemoveCaptionClick(Office.IRibbonControl control)
        {
            removeCaptionsEnabled = false;
            RefreshRibbonControl("removeCaptions");
            NotesToCaptions.RemoveCaptionsFromSelectedSlides();
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

        public void FitToWidthClick(Office.IRibbonControl control)
        {
            var selectedShape = PowerPointPresentation.CurrentSelection.ShapeRange[1];
            FitToSlide.FitToWidth(selectedShape);
        }

        public void FitToHeightClick(Office.IRibbonControl control)
        {
            var selectedShape = PowerPointPresentation.CurrentSelection.ShapeRange[1];
            FitToSlide.FitToHeight(selectedShape);
        }

        public System.Drawing.Bitmap GetFitToWidthImage(Office.IRibbonControl control)
        {
            return FitToSlide.GetFitToWidthImage(control);
        }

        public System.Drawing.Bitmap GetFitToHeightImage(Office.IRibbonControl control)
        {
            return FitToSlide.GetFitToHeightImage(control);
        }

        #endregion

        #region feature: Crop to Shape

        public void CropShapeButtonClick(Office.IRibbonControl control)
        {
            var selection = PowerPointPresentation.CurrentSelection;
            CropToShape.Crop(selection);
        }

        public System.Drawing.Bitmap GetCutOutShapeMenuImage(Office.IRibbonControl control)
        {
            return CropToShape.GetCutOutShapeMenuImage(control);
        }

        #endregion

        #region feature: Convert to Picture

        public void ConvertToPictureButtonClick(Office.IRibbonControl control)
        {
            var selection = PowerPointPresentation.CurrentSelection;
            ConvertToPicture.Convert(selection);
        }

        public System.Drawing.Bitmap GetConvertToPicMenuImage(Office.IRibbonControl control)
        {
            return ConvertToPicture.GetConvertToPicMenuImage(control);
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
