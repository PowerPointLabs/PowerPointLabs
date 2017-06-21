using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Diagnostics.CodeAnalysis;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;

using PowerPointLabs.ActionFramework.Common.Factory;
using PowerPointLabs.ActionFramework.Common.Log;
using PowerPointLabs.CropLab;
using PowerPointLabs.DataSources;
using PowerPointLabs.DrawingsLab;
using PowerPointLabs.HighlightLab;
using PowerPointLabs.Models;
using PowerPointLabs.PictureSlidesLab.View;
using PowerPointLabs.Views;

using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;


// Follow these steps to enable the Ribbon (XML) item:

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
    [SuppressMessage("Microsoft.StyleCop.CSharp.OrderingRules", "SA1202:ElementsMustBeOrderedByAccess", Justification = "Migration to Action Framework")]
    public class Ribbon1 : Office.IRibbonExtensibility
    {
        #region Action Framework Factory
        private ActionHandlerFactory ActionHandlerFactory { get; set; }

        private EnabledHandlerFactory EnabledHandlerFactory { get; set; }

        private LabelHandlerFactory LabelHandlerFactory { get; set; }

        private ImageHandlerFactory ImageHandlerFactory { get; set; }

        private SupertipHandlerFactory SupertipHandlerFactory { get; set; }

        private ContentHandlerFactory ContentHandlerFactory { get; set; }

        private PressedHandlerFactory PressedHandlerFactory { get; set; }

        private CheckBoxActionHandlerFactory CheckBoxActionHandlerFactory { get; set; }
        #endregion

        #region Action Framework entry point

        public void OnAction(Office.IRibbonControl control)
        {
            var actionHandler = ActionHandlerFactory.CreateInstance(control.Id, control.Tag);
            actionHandler.Execute(control.Id);
        }

        public bool GetEnabled(Office.IRibbonControl control)
        {
            var enabledHandler = EnabledHandlerFactory.CreateInstance(control.Id, control.Tag);
            return enabledHandler.Get(control.Id);
        }

        public string GetLabel(Office.IRibbonControl control)
        {
            var labelHandler = LabelHandlerFactory.CreateInstance(control.Id, control.Tag);
            return labelHandler.Get(control.Id);
        }

        public string GetSupertip(Office.IRibbonControl control)
        {
            var supertipHandler = SupertipHandlerFactory.CreateInstance(control.Id, control.Tag);
            return supertipHandler.Get(control.Id);
        }

        public Bitmap GetImage(Office.IRibbonControl control)
        {
            var imageHandler = ImageHandlerFactory.CreateInstance(control.Id, control.Tag);
            return imageHandler.Get(control.Id);
        }

        public string GetContent(Office.IRibbonControl control)
        {
            var contentHandler = ContentHandlerFactory.CreateInstance(control.Id, control.Tag);
            return contentHandler.Get(control.Id);
        }

        public bool GetPressed(Office.IRibbonControl control)
        {
            var pressedHandler = PressedHandlerFactory.CreateInstance(control.Id, control.Tag);
            return pressedHandler.Get(control.Id);
        }

        public void OnCheckBoxAction(Office.IRibbonControl control, bool pressed)
        {
            var checkBoxActionHandler = CheckBoxActionHandlerFactory.CreateInstance(control.Id, control.Tag);
            checkBoxActionHandler.Execute(control.Id, pressed);
        }

        #endregion

        #region Deprecated. Please only use Action Framework to support the feature.

#pragma warning disable 0618
        private Office.IRibbonUI _ribbon;

        public bool FrameAnimationChecked = false;
        public bool BackgroundZoomChecked = true;
        public bool MultiSlideZoomChecked = true;
        public bool SpotlightDelete = true;
        public float DefaultDuration = 0.5f;

        public bool SpotlightEnabled = false;
        public bool InSlideEnabled = false;
        public bool ZoomButtonEnabled = false;
        public bool HighlightBulletsEnabled = true;
        public bool AddAutoMotionEnabled = true;
        public bool ReloadAutoMotionEnabled = true;
        public bool ReloadSpotlight = true;
        public bool RemoveCaptionsEnabled = true;
        public bool RemoveAudioEnabled = true;

        public bool HighlightTextFragmentsEnabled = true;

        public bool EmbedAudioVisible = true;
        public bool RecorderPaneVisible = false;

        private bool _previewCurrentSlide;

        private List<string> _voiceNames;

        private int _voiceSelected;

        #region IRibbonExtensibility Members

        public string GetCustomUI(string ribbonId)
        {
            return GetResourceText("PowerPointLabs.Ribbon1.xml");
        }

        #endregion

        #region Ribbon Callbacks
        //Create callback methods here. For more information about adding callback methods, select the Ribbon XML item in Solution Explorer and then press F1

        public void RibbonLoad(Office.IRibbonUI ribbonUi)
        {
            ActionHandlerFactory = new ActionHandlerFactory();
            EnabledHandlerFactory = new EnabledHandlerFactory();
            LabelHandlerFactory = new LabelHandlerFactory();
            SupertipHandlerFactory = new SupertipHandlerFactory();
            ImageHandlerFactory = new ImageHandlerFactory();
            ContentHandlerFactory = new ContentHandlerFactory();
            PressedHandlerFactory = new PressedHandlerFactory();
            CheckBoxActionHandlerFactory = new CheckBoxActionHandlerFactory();

            _ribbon = ribbonUi;

            SetVoicesFromInstalledOptions();
            SetCoreVoicesToSelections();
        }

        public void RefreshRibbonControl(String controlId)
        {
            try
            {
                _ribbon.InvalidateControl(controlId);
            }
            catch (Exception e)
            {
                Logger.LogException(e, "RefreshRibbonControl");
                throw;
            }
        }

        private void SetVoicesFromInstalledOptions()
        {
            var installedVoices = NotesToAudio.GetVoices().ToList();
            _voiceNames = installedVoices;
        }

        public void SpotlightBtnClick(Office.IRibbonControl control)
        {
            try
            {
                if (Globals.ThisAddIn.Application.ActiveWindow.Selection.Type != PowerPoint.PpSelectionType.ppSelectionShapes)
                {
                    return;
                }

                Globals.ThisAddIn.Application.StartNewUndoEntry();

                Spotlight.AddSpotlightEffect();
            }
            catch (Exception e)
            {
                Logger.LogException(e, "SpotlightBtnClick");
                throw;
            }
        }

        # region Supertipsa
        public string GetAddSpotlightButtonSupertip(Office.IRibbonControl control)
        {
            return TextCollection.AddSpotlightButtonSupertip;
        }

        public string GetSpotlightPropertiesButtonSupertip(Office.IRibbonControl control)
        {
            return TextCollection.SpotlightPropertiesButtonSupertip;
        }

        public string GetEffectsLabSupertip(Office.IRibbonControl control)
        {
            return TextCollection.EffectsLabMenuSupertip;
        }
        public string GetEffectsLabMakeTransparentSupertip(Office.IRibbonControl control)
        {
            return TextCollection.EffectsLabMakeTransparentSupertip;
        }
        public string GetEffectsLabMagnifyGlassSupertip(Office.IRibbonControl control)
        {
            return TextCollection.EffectsLabMagnifyGlassSupertip;
        }
        public string GetEffectsLabBlurRemainderSupertip(Office.IRibbonControl control)
        {
            return TextCollection.EffectsLabBlurRemainderSupertip;
        }
        public string GetEffectsLabColorizeRemainderSupertip(Office.IRibbonControl control)
        {
            return TextCollection.EffectsLabColorizeRemainderSupertip;
        }
        public string GetEffectsLabBlurBackgroundSupertip(Office.IRibbonControl control)
        {
            return TextCollection.EffectsLabBlurBackgroundSupertip;
        }
        public string GetEffectsLabColorizeBackgroundSupertip(Office.IRibbonControl control)
        {
            return TextCollection.EffectsLabColorizeBackgroundSupertip;
        }

        public string GetDrawingsLabButtonSupertip(Office.IRibbonControl control)
        {
            return TextCollection.DrawingsLabButtonSupertip;
        }

        public string GetResizeLabButtonSupertip(Office.IRibbonControl control)
        {
            return TextCollection.ResizeLabMenuSupertip;
        }

        public string GetHelpButtonSupertip(Office.IRibbonControl control)
        {
            return TextCollection.UserGuideButtonSupertip;
        }
        public string GetFeedbackButtonSupertip(Office.IRibbonControl control)
        {
            return TextCollection.FeedbackButtonSupertip;
        }
        public string GetAboutButtonSupertip(Office.IRibbonControl control)
        {
            return TextCollection.AboutButtonSupertip;
        }
        public string GetPositionsLabSupertip(Office.IRibbonControl control)
        {
            return TextCollection.PositionsLabMenuSupertip;
        }
        # endregion

        # region Button Labels
        public string GetPowerPointLabsAddInsTabLabel(Office.IRibbonControl control)
        {
            return TextCollection.PowerPointLabsAddInsTabLabel;
        }

        public string GetCombineShapesLabel(Office.IRibbonControl control)
        {
            return TextCollection.CombineShapesLabel;
        }
        
        public string GetSpotlightMenuLabel(Office.IRibbonControl control)
        {
            return TextCollection.SpotlightMenuLabel;
        }
        public string GetAddSpotlightButtonLabel(Office.IRibbonControl control)
        {
            return TextCollection.AddSpotlightButtonLabel;
        }
        public string GetReloadSpotlightButtonLabel(Office.IRibbonControl control)
        {
            return TextCollection.ReloadSpotlightButtonLabel;
        }
        
        public string GetSpotlightPropertiesButtonLabel(Office.IRibbonControl control)
        {
            return TextCollection.SpotlightPropertiesButtonLabel;
        }

        public string GetEffectsLabButtonLabel(Office.IRibbonControl control)
        {
            return TextCollection.EffectsLabButtonLabel;
        }
        public string GetEffectsLabMakeTransparentButtonLabel(Office.IRibbonControl control)
        {
            return TextCollection.EffectsLabMakeTransparentButtonLabel;
        }
        public string GetEffectsLabMagnifyGlassButtonLabel(Office.IRibbonControl control)
        {
            return TextCollection.EffectsLabMagnifyGlassButtonLabel;
        }
        public string GetEffectsLabBlurRemainderButtonLabel(Office.IRibbonControl control)
        {
            return TextCollection.EffectsLabBlurRemainderButtonLabel;
        }
        public string GetEffectsLabBlurBackgroundButtonLabel(Office.IRibbonControl control)
        {
            return TextCollection.EffectsLabBlurBackgroundButtonLabel;
        }
        public string GetEffectsLabRecolorRemainderButtonLabel(Office.IRibbonControl control)
        {
            return TextCollection.EffectsLabRecolorRemainderButtonLabel;
        }
        public string GetEffectsLabRecolorBackgroundButtonLabel(Office.IRibbonControl control)
        {
            return TextCollection.EffectsLabRecolorBackgroundButtonLabel;
        }

        public string GetAgendaLabButtonLabel(Office.IRibbonControl control)
        {
            return TextCollection.AgendaLabButtonLabel;
        }
        public string GetAgendaLabBulletPointButtonLabel(Office.IRibbonControl control)
        {
            return TextCollection.AgendaLabBulletPointButtonLabel;
        }
        public string GetAgendaLabAgendaSettingsButtonLabel(Office.IRibbonControl control)
        {
            return TextCollection.AgendaLabAgendaSettingsButtonLabel;
        }
        public string GetAgendaLabBulletAgendaSettingsButtonLabel(Office.IRibbonControl control)
        {
            return TextCollection.AgendaLabBulletAgendaSettingsButtonLabel;
        }

        public string GetDrawingsLabButtonLabel(Office.IRibbonControl control)
        {
            return TextCollection.DrawingsLabButtonLabel;
        }

        public string GetPositionsLabButtonLabel(Office.IRibbonControl control)
        {
            return TextCollection.PositionsLabButtonLabel;
        }

        public string GetHelpButtonLabel(Office.IRibbonControl control)
        {
            return TextCollection.UserGuideButtonLabel;
        }
        public string GetFeedbackButtonLabel(Office.IRibbonControl control)
        {
            return TextCollection.FeedbackButtonLabel;
        }
        public string GetAboutButtonLabel(Office.IRibbonControl control)
        {
            return TextCollection.AboutButtonLabel;
        }

        public string GetContextSpeakSelectedTextLabel(Office.IRibbonControl control)
        {
            return TextCollection.ContextSpeakSelectedTextLabel;
        }
        public string GetContextAddCurrentSlideLabel(Office.IRibbonControl control)
        {
            return TextCollection.ContextAddCurrentSlideLabel;
        }
        public string GetContextReplaceAudioLabel(Office.IRibbonControl control)
        {
            return TextCollection.ContextReplaceAudioLabel;
        }
        public string GetPowerPointLabsMenuLabel(Office.IRibbonControl control)
        {
            return TextCollection.PowerPointLabsMenuLabel;
        }
        # endregion

        //Button Click Callbacks
        public void AboutButtonClick(Office.IRibbonControl control)
        {
            MessageBox.Show(TextCollection.AboutInfo, TextCollection.AboutInfoTitle);
        }
        public void HelpButtonClick(Office.IRibbonControl control)
        {
            try
            {
                Process.Start(TextCollection.HelpDocumentUrl);
            }
            catch (Exception e)
            {
                Logger.LogException(e, "HelpButtonClick");
                throw;
            }
        }
        public void TutorialButtonClick(Office.IRibbonControl control)
        {
            string sourceFile = "";
            switch (Properties.Settings.Default.ReleaseType)
            {
                case "dev":
                    sourceFile = Properties.Settings.Default.DevAddr + TextCollection.QuickTutorialFileName;
                    break;
                case "release":
                    sourceFile = Properties.Settings.Default.ReleaseAddr + TextCollection.QuickTutorialFileName;
                    break;
            }

            try
            {
                if (sourceFile != "")
                {
                    Process.Start("POWERPNT", sourceFile);
                }
            }
            catch
            {
                Logger.Log("TutorialButtonClick: Failed to open tutorial file!", ActionFramework.Common.Logger.LogType.Error);
            }
        }
        public void FeedbackButtonClick(Office.IRibbonControl control)
        {
            try
            {
                Process.Start(TextCollection.FeedbackUrl);
            }
            catch (Exception e)
            {
                Logger.LogException(e, "FeedbackButtonClick");
                throw;
            }
        }

        # region Icon Getters
        public Bitmap GetSpotlightImage(Office.IRibbonControl control)
        {
            try
            {
                return new Bitmap(Properties.Resources.Spotlight);
            }
            catch (Exception e)
            {
                Logger.LogException(e, "GetSpotlightImage");
                throw;
            }
        }
        public Bitmap GetReloadSpotlightImage(Office.IRibbonControl control)
        {
            try
            {
                return new Bitmap(Properties.Resources.ReloadSpotlight);
            }
            catch (Exception e)
            {
                Logger.LogException(e, "GetReloadSpotlightImage");
                throw;
            }
        }

        public Bitmap GetEffectsLabImage(Office.IRibbonControl control)
        {
            try
            {
                return new System.Drawing.Bitmap(Properties.Resources.EffectsLab);
            }
            catch (Exception e)
            {
                Logger.LogException(e, "GetEffectsLabImage");
                throw;
            }
        }
        public Bitmap GetMakeTransparentImage(Office.IRibbonControl control)
        {
            try
            {
                return new System.Drawing.Bitmap(Properties.Resources.MakeTransparent);
            }
            catch (Exception e)
            {
                Logger.LogException(e, "GetMakeTransparentImage");
                throw;
            }
        }
        public Bitmap GetMagnifyImage(Office.IRibbonControl control)
        {
            try
            {
                return new System.Drawing.Bitmap(Properties.Resources.Magnify);
            }
            catch (Exception e)
            {
                Logger.LogException(e, "GetMagnifyImage");
                throw;
            }
        }
        public Bitmap GetBlurRemainderImage(Office.IRibbonControl control)
        {
            try
            {
                return new System.Drawing.Bitmap(Properties.Resources.BlurRemainder);
            }
            catch (Exception e)
            {
                Logger.LogException(e, "GetBlurRemainderImage");
                throw;
            }
        }
        public Bitmap GetRecolorRemainderImage(Office.IRibbonControl control)
        {
            try
            {
                return new System.Drawing.Bitmap(Properties.Resources.RecolorRemainder);
            }
            catch (Exception e)
            {
                Logger.LogException(e, "GetRecolorRemainderImage");
                throw;
            }
        }

        public Bitmap GetAgendaSettingsImage(Office.IRibbonControl control)
        {
            try
            {
                return new Bitmap(Properties.Resources.AgendaSettings);
            }
            catch (Exception e)
            {
                Logger.LogException(e, "GetAgendaSettingsImage");
                throw;
            }
        }
        public Bitmap GetDrawingsLabImage(Office.IRibbonControl control)
        {
            try
            {
                return new System.Drawing.Bitmap(Properties.Resources.DrawingLab);
            }
            catch (Exception e)
            {
                Logger.LogException(e, "GetDrawingsLabImage");
                throw;
            }
        }

        public Bitmap GetAboutImage(Office.IRibbonControl control)
        {
            try
            {
                return new Bitmap(Properties.Resources.About);
            }
            catch (Exception e)
            {
                Logger.LogException(e, "GetAboutImage");
                throw;
            }
        }
        public Bitmap GetHelpImage(Office.IRibbonControl control)
        {
            try
            {
                return new Bitmap(Properties.Resources.UserGuide);
            }
            catch (Exception e)
            {
                Logger.LogException(e, "GetHelpImage");
                throw;
            }
        }
        public Bitmap GetFeedbackImage(Office.IRibbonControl control)
        {
            try
            {
                return new Bitmap(Properties.Resources.Feedback);
            }
            catch (Exception e)
            {
                Logger.LogException(e, "GetFeedbackImage");
                throw;
            }
        }

        public Bitmap GetPreviewNarrationContextImage(Office.IRibbonControl control)
        {
            try
            {
                return new Bitmap(Properties.Resources.SpeakTextContext);
            }
            catch (Exception e)
            {
                Logger.LogException(e, "GetPreviewNarrationContextImage");
                throw;
            }
        }

        public Bitmap GetAddSpotlightContextImage(Office.IRibbonControl control)
        {
            try
            {
                return new Bitmap(Properties.Resources.AddSpotlightContext);
            }
            catch (Exception e)
            {
                Logger.LogException(e, "GetAddSpotlightContextImage");
                throw;
            }
        }
        # endregion

        public void ZoomStyleChanged(Office.IRibbonControl control, bool pressed)
        {
            try
            {
                BackgroundZoomChecked = pressed;
            }
            catch (Exception e)
            {
                Logger.LogException(e, "ZoomStyleChanged");
                throw;
            }
        }
        public bool ZoomStyleGetPressed(Office.IRibbonControl control)
        {
            try
            {
                return BackgroundZoomChecked;
            }
            catch (Exception e)
            {
                Logger.LogException(e, "ZoomStyleGetPressed");
                throw;
            }
        }

        # region Control Enable
        public bool OnGetEnabledSpotlight(Office.IRibbonControl control)
        {
            return SpotlightEnabled;
        }
        public bool OnGetEnabledReloadSpotlight(Office.IRibbonControl control)
        {
            return ReloadSpotlight;
        }
        public bool OnGetEnabledAddAutoMotion(Office.IRibbonControl control)
        {
            return AddAutoMotionEnabled;
        }
        public bool OnGetEnabledAddInSlide(Office.IRibbonControl control)
        {
            return InSlideEnabled;
        }
        # endregion

        //Edit Name Callbacks
        public void ShapeNameEdited(String newName)
        {
            try
            {
                Globals.ThisAddIn.Application.StartNewUndoEntry();

                PowerPoint.Shape selectedShape = Globals.ThisAddIn.Application.ActiveWindow.Selection.ShapeRange[1];
                selectedShape.Name = newName;
            }
            catch (Exception e)
            {
                Logger.LogException(e, "ShapeNameEdited");
                throw;
            }
        }

        public void AnimationLabDialogButtonPressed(Office.IRibbonControl control)
        {
            try
            {
                var dialog = new AnimationLabDialogBox(DefaultDuration, FrameAnimationChecked);
                dialog.SettingsHandler += AnimationPropertiesEdited;
                dialog.ShowDialog();
            }
            catch (Exception e)
            {
                Logger.LogException(e, "AnimationLabDialogButtonPressed");
                throw;
            }
        }

        public void AnimationPropertiesEdited(float newDuration, bool newFrameChecked)
        {
            try
            {
                DefaultDuration = newDuration;
                FrameAnimationChecked = newFrameChecked;
                AnimateInSlide.defaultDuration = newDuration;
                AnimateInSlide.frameAnimationChecked = newFrameChecked;
                AutoAnimate.defaultDuration = newDuration;
                AutoAnimate.frameAnimationChecked = newFrameChecked;
            }
            catch (Exception e)
            {
                Logger.LogException(e, "AnimationPropertiesEdited");
                throw;
            }
        }

        public void ZoomLabDialogButtonPressed(Office.IRibbonControl control)
        {
            try
            {
                var dialog = new ZoomLabDialogBox(BackgroundZoomChecked, MultiSlideZoomChecked);
                dialog.SettingsHandler += ZoomPropertiesEdited;
                dialog.ShowDialog();
            }
            catch (Exception e)
            {
                Logger.LogException(e, "ZoomLabDialogButtonPressed");
                throw;
            }
        }

        public void ZoomPropertiesEdited(bool backgroundChecked, bool multiSlideChecked)
        {
            try
            {
                BackgroundZoomChecked = backgroundChecked;
                MultiSlideZoomChecked = multiSlideChecked;
                AutoZoom.backgroundZoomChecked = backgroundChecked;
                ZoomToArea.backgroundZoomChecked = backgroundChecked;
                ZoomToArea.multiSlideZoomChecked = multiSlideChecked;
            }
            catch (Exception e)
            {
                Logger.LogException(e, "ZoomPropertiesEdited");
                throw;
            }
        }

        public void SpotlightDialogButtonPressed(Office.IRibbonControl control)
        {
            try
            {
                var dialog = new SpotlightDialogBox(Spotlight.defaultTransparency, Spotlight.defaultSoftEdges,
                    Spotlight.defaultColor);
                dialog.SettingsHandler += SpotlightPropertiesEdited;
                dialog.ShowDialog();
            }
            catch (Exception e)
            {
                Logger.LogException(e, "SpotlightDialogButtonPressed");
                throw;
            }
        }

        public void SpotlightPropertiesEdited(float newTransparency, float newSoftEdge, Color newColor)
        {
            try
            {
                Spotlight.defaultTransparency = newTransparency;
                Spotlight.defaultSoftEdges = newSoftEdge;
                Spotlight.defaultColor = newColor;
            }
            catch (Exception e)
            {
                Logger.LogException(e, "SpotlightPropertiesEdited");
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
                HighlightTextFragments.backgroundColor = newBackgroundColor;
            }
            catch (Exception e)
            {
                Logger.LogException(e, "HighlightBulletsPropertiesEdited");
                throw;
            }
        }
        public void HighlightBulletsDialogBoxPressed(Office.IRibbonControl control)
        {
            try
            {
                var dialog = new HighlightBulletsDialogBox(HighlightBulletsText.highlightColor, HighlightBulletsText.defaultColor, HighlightBulletsBackground.backgroundColor);
                dialog.SettingsHandler += HighlightBulletsPropertiesEdited;
                dialog.ShowDialog();
            }
            catch (Exception e)
            {
                Logger.LogException(e, "HighlightBulletsDialogBoxPressed");
                throw;
            }
        }

        public bool GetEmbedAudioVisiblity(Office.IRibbonControl control)
        {
            return EmbedAudioVisible;
        }

        public bool IsValidPresentation(PowerPoint.Presentation pres)
        {
            if (!Globals.ThisAddIn.VerifyVersion(pres))
            {
                MessageBox.Show(TextCollection.VersionNotCompatibleErrorMsg);
                return false;
            }

            return true;
        }

        public void PreviewAnimationsIfChecked()
        {
            if (_previewCurrentSlide)
            {
                NotesToAudio.PreviewAnimations();
            }
        }

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

        public Bitmap GetContextMenuImage(Office.IRibbonControl control)
        {
            try
            {
                return new Bitmap(Properties.Resources.PptlabsContextMenu);
            }
            catch (Exception e)
            {
                Logger.LogException(e, "GetContextMenuImage");
                throw;
            }
        }

        #endregion

        #region Feature: Picture Slides Lab

        public PictureSlidesLabWindow PictureSlidesLabWindow { get; set; }

        #endregion

        # region Feature: Combine Shapes
        public bool GetVisibilityForCombineShapes(Office.IRibbonControl control)
        {
            const string officeVersion2010 = "14.0";
            return Globals.ThisAddIn.Application.Version == officeVersion2010;
        }
        # endregion

        # region Feature: Narrations Lab
        public void NarrationsLabDialogButtonPressed(Office.IRibbonControl control)
        {
            try
            {
                var dialog = new NarrationsLabDialogBox(_voiceSelected, _voiceNames,
                    _previewCurrentSlide);
                dialog.SettingsHandler += NarrationsLabSettingsChanged;
                dialog.ShowDialog();
            }
            catch (Exception e)
            {
                Logger.LogException(e, "NarrationsLabDialogButtonPressed");
                throw;
            }
        }

        public void NarrationsLabSettingsChanged(String voiceName, bool previewCurrentSlide)
        {
            _previewCurrentSlide = previewCurrentSlide;
            if (!String.IsNullOrWhiteSpace(voiceName))
            {
                NotesToAudio.SetDefaultVoice(voiceName);
                _voiceSelected = _voiceNames.IndexOf(voiceName);
            }
        }

        public void SpeakSelectedTextClick(Office.IRibbonControl control)
        {
            NotesToAudio.SpeakSelectedText();
        }
        # endregion

        # region Feature: Effects Lab
        public void MagnifyGlassEffectClick(Office.IRibbonControl control)
        {
            Globals.ThisAddIn.Application.StartNewUndoEntry();

            var selection = PowerPointCurrentPresentationInfo.CurrentSelection;

            PowerPoint.ShapeRange shapeRange;

            try
            {
                shapeRange = selection.ShapeRange;
            }
            catch (Exception)
            {
                MessageBox.Show("Please select an area to magnify.");

                return;
            }

            if (shapeRange.Count > 1 || shapeRange[1].Type == Office.MsoShapeType.msoGroup)
            {
                MessageBox.Show("Only one magnify area is allowed.");

                return;
            }

            try
            {
                var croppedShape = CropToShape.Crop(PowerPointCurrentPresentationInfo.CurrentSlide, selection, isInPlace: true, handleError: false);

                MagnifyGlassEffect(croppedShape, 1.4f);
            }
            catch (Exception e)
            {
                var errorMessage = e.Message;
                errorMessage = errorMessage.Replace("Crop To Shape", "Magnify");

                MessageBox.Show(errorMessage);
            }
        }

        public void BlurRemainderEffectClick(int percentage)
        {
            Globals.ThisAddIn.Application.StartNewUndoEntry();

            var effectSlide = GenerateEffectSlide(true);

            if (effectSlide == null)
            {
                return;
            }

            effectSlide.BlurBackground(percentage, EffectsLab.EffectsLabBlurSelected.IsTintRemainder);
            effectSlide.GetNativeSlide().Select();
        }

        public void GreyScaleRemainderEffectClick(Office.IRibbonControl control)
        {
            Globals.ThisAddIn.Application.StartNewUndoEntry();

            var effectSlide = GenerateEffectSlide(true);

            if (effectSlide == null)
            {
                return;
            }

            effectSlide.GreyScaleBackground();
            effectSlide.GetNativeSlide().Select();
        }

        public void BlackWhiteRemainderEffectClick(Office.IRibbonControl control)
        {
            Globals.ThisAddIn.Application.StartNewUndoEntry();

            var effectSlide = GenerateEffectSlide(true);

            if (effectSlide == null)
            {
                return;
            }

            effectSlide.BlackWhiteBackground();
            effectSlide.GetNativeSlide().Select();
        }

        public void GothamRemainderEffectClick(Office.IRibbonControl control)
        {
            Globals.ThisAddIn.Application.StartNewUndoEntry();

            var effectSlide = GenerateEffectSlide(true);

            if (effectSlide == null)
            {
                return;
            }

            effectSlide.GothamBackground();
            effectSlide.GetNativeSlide().Select();
        }

        public void SepiaRemainderEffectClick(Office.IRibbonControl control)
        {
            Globals.ThisAddIn.Application.StartNewUndoEntry();

            var effectSlide = GenerateEffectSlide(true);

            if (effectSlide == null)
            {
                return;
            }

            effectSlide.SepiaBackground();
            effectSlide.GetNativeSlide().Select();
        }

        public void BlurBackgroundEffectClick(int percentage)
        {
            Globals.ThisAddIn.Application.StartNewUndoEntry();

            var effectSlide = GenerateEffectSlide(false);

            if (effectSlide == null)
            {
                return;
            }

            effectSlide.BlurBackground(percentage, EffectsLab.EffectsLabBlurSelected.IsTintBackground);
            effectSlide.GetNativeSlide().Select();
        }

        public void GreyScaleBackgroundEffectClick(Office.IRibbonControl control)
        {
            Globals.ThisAddIn.Application.StartNewUndoEntry();

            var effectSlide = GenerateEffectSlide(false);

            if (effectSlide == null)
            {
                return;
            }

            effectSlide.GreyScaleBackground();
            effectSlide.GetNativeSlide().Select();
        }

        public void BlackWhiteBackgroundEffectClick(Office.IRibbonControl control)
        {
            Globals.ThisAddIn.Application.StartNewUndoEntry();

            var effectSlide = GenerateEffectSlide(false);

            if (effectSlide == null)
            {
                return;
            }

            effectSlide.BlackWhiteBackground();
            effectSlide.GetNativeSlide().Select();
        }

        public void GothamBackgroundEffectClick(Office.IRibbonControl control)
        {
            Globals.ThisAddIn.Application.StartNewUndoEntry();

            var effectSlide = GenerateEffectSlide(false);

            if (effectSlide == null)
            {
                return;
            }

            effectSlide.GothamBackground();
            effectSlide.GetNativeSlide().Select();
        }

        public void SepiaBackgroundEffectClick(Office.IRibbonControl control)
        {
            Globals.ThisAddIn.Application.StartNewUndoEntry();

            var effectSlide = GenerateEffectSlide(false);

            if (effectSlide == null)
            {
                return;
            }

            effectSlide.SepiaBackground();
            effectSlide.GetNativeSlide().Select();
        }

        public void TransparentEffectClick(Office.IRibbonControl control)
        {
            Globals.ThisAddIn.Application.StartNewUndoEntry();

            var selection = PowerPointCurrentPresentationInfo.CurrentSelection;

            if (selection.Type != PowerPoint.PpSelectionType.ppSelectionShapes)
            {
                MessageBox.Show("Please select at least 1 shape");
                return;
            }

            TransparentEffect(selection.ShapeRange);
        }

        private void MagnifyGlassEffect(PowerPoint.Shape shape, float ratio)
        {
            var delta = 0.5f * (ratio - 1);

            shape.Left -= delta * shape.Width;
            shape.Top -= delta * shape.Height;

            shape.Width *= ratio;
            shape.Height *= ratio;

            shape.Shadow.Visible = Office.MsoTriState.msoTrue;
            shape.Shadow.Style = Office.MsoShadowStyle.msoShadowStyleOuterShadow;
            shape.Shadow.Transparency = 0.6f;
            shape.Shadow.Size = 102f;
            shape.Shadow.Blur = 5;
            shape.Shadow.OffsetX = 0;
            shape.Shadow.OffsetY = 2f;

            shape.ThreeD.BevelTopType = Office.MsoBevelType.msoBevelCircle;
            shape.ThreeD.BevelTopInset = 15;
            shape.ThreeD.BevelTopDepth = 3;
            shape.ThreeD.BevelBottomType = Office.MsoBevelType.msoBevelNone;
            shape.ThreeD.PresetLighting = Office.MsoLightRigType.msoLightRigBalanced;
            shape.ThreeD.LightAngle = 145;

            shape.LockAspectRatio = Office.MsoTriState.msoTrue;
        }

        private PowerPointBgEffectSlide GenerateEffectSlide(bool generateOnRemainder)
        {
            var curSlide = PowerPointCurrentPresentationInfo.CurrentSlide;
            var selection = PowerPointCurrentPresentationInfo.CurrentSelection;
            PowerPointSlide dupSlide = null;

            try
            {
                var shapeRange = selection.ShapeRange;

                if (shapeRange.Count != 0)
                {
                    dupSlide = curSlide.Duplicate();
                }

                shapeRange.Cut();

                var effectSlide = PowerPointBgEffectSlide.BgEffectFactory(curSlide.GetNativeSlide(), generateOnRemainder);

                if (dupSlide != null)
                {
                    if (generateOnRemainder)
                    {
                        dupSlide.Delete();
                    }
                    else
                    {
                        dupSlide.MoveTo(curSlide.Index);
                        curSlide.Delete();
                    }
                }

                PowerPointPresentation.Current.AddAckSlide();

                return effectSlide;
            }
            catch (InvalidOperationException e)
            {
                MessageBox.Show(e.Message);
                return null;
            }
            catch (COMException)
            {
                if (dupSlide != null)
                {
                    dupSlide.Delete();
                }

                MessageBox.Show("Please select at least 1 shape");
                return null;
            }
            catch (Exception e)
            {
                if (dupSlide != null)
                {
                    dupSlide.Delete();
                }

                ErrorDialogWrapper.ShowDialog("Error", e.Message, e);
                return null;
            }
        }

        private void TransparentEffect(PowerPoint.ShapeRange shapeRange)
        {
            foreach (PowerPoint.Shape shape in shapeRange)
            {
                if (shape.Type == Office.MsoShapeType.msoGroup)
                {
                    var subShapeRange = shape.Ungroup();
                    TransparentEffect(subShapeRange);
                    subShapeRange.Group();
                }
                else if (shape.Type == Office.MsoShapeType.msoPlaceholder)
                {
                    PlaceholderTransparencyHandler(shape);
                }
                else if (shape.Type == Office.MsoShapeType.msoPicture)
                {
                    PictureTransparencyHandler(shape);
                }
                else if (shape.Type == Office.MsoShapeType.msoLine)
                {
                    LineTransparencyHandler(shape);
                }
                else if (IsTransparentableShape(shape))
                {
                    ShapeTransparencyHandler(shape);
                }
            }
        }

        private bool IsTransparentableShape(PowerPoint.Shape shape)
        {
            return shape.Type == Office.MsoShapeType.msoAutoShape ||
                   shape.Type == Office.MsoShapeType.msoFreeform;
        }

        private void PictureTransparencyHandler(PowerPoint.Shape picture)
        {
            var rotation = picture.Rotation;

            picture.Rotation = 0;

            var tempPicPath = Path.Combine(Path.GetTempPath(), "tempPic.png");

            Utils.Graphics.ExportShape(picture, tempPicPath);

            var shapeHolder =
                PowerPointCurrentPresentationInfo.CurrentSlide.Shapes.AddShape(
                    Office.MsoAutoShapeType.msoShapeRectangle,
                    picture.Left,
                    picture.Top,
                    picture.Width,
                    picture.Height);

            var oriZOrder = picture.ZOrderPosition;

            picture.Delete();

            // move shape holder to original z-order
            while (shapeHolder.ZOrderPosition > oriZOrder)
            {
                shapeHolder.ZOrder(Office.MsoZOrderCmd.msoSendBackward);
            }

            shapeHolder.Line.Visible = Office.MsoTriState.msoFalse;
            shapeHolder.Fill.UserPicture(tempPicPath);
            shapeHolder.Fill.Transparency = 0.5f;

            shapeHolder.Rotation = rotation;

            File.Delete(tempPicPath);
        }

        private void PlaceholderTransparencyHandler(PowerPoint.Shape picture)
        {
            PictureTransparencyHandler(picture);
        }

        private void LineTransparencyHandler(PowerPoint.Shape shape)
        {
            shape.Line.Transparency = 0.5f;
        }

        private void ShapeTransparencyHandler(PowerPoint.Shape shape)
        {
            shape.Fill.Transparency = 0.5f;
            shape.Line.Transparency = 0.5f;
        }
        # endregion

        #region Feature: Drawing Lab
        internal DrawingLabData DrawingLabData { get; set; }
        internal DrawingsLabMain DrawingLab { get; set; }

        public void DrawingsLabButtonClick(Office.IRibbonControl control)
        {
            try
            {
                if (DrawingLabData == null)
                {
                    DrawingLabData = new DrawingLabData();
                    DrawingLab = new DrawingsLabMain(DrawingLabData);
                }

                Globals.ThisAddIn.RegisterDrawingsPane(PowerPointPresentation.Current.Presentation);

                var drawingsPane = Globals.ThisAddIn.GetActivePane(typeof(DrawingsPane));
                ((DrawingsPane)drawingsPane.Control).drawingsPaneWPF.TryInitialise(DrawingLabData, DrawingLab);

                // if currently the pane is hidden, show the pane
                if (!drawingsPane.Visible)
                {
                    // fire the pane visble change event
                    drawingsPane.Visible = true;
                }
                else
                {
                    drawingsPane.Visible = false;
                }
            }
            catch (Exception e)
            {
                ErrorDialogWrapper.ShowDialog("Error in drawing lab", e.Message, e);
                Logger.LogException(e, "DrawingsLabButtonClicked");
                throw;
            }
        }
        #endregion

        // TODO: Add the image for the icon on the ribbon bar
        //public Bitmap GetPositionsLabImage(Office.IRibbonControl control)
        //{
        //    try
        //    {
        //        return new Bitmap(Properties.Resources.PositionsLab);
        //    }
        //    catch (Exception e)
        //    {
        //        PowerPointLabsGlobals.LogException(e, "GetPositionsLabImage");
        //        throw;
        //    }
        //}

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
        #endregion
    }
}
