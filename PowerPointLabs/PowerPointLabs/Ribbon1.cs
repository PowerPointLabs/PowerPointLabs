using System;
using System.Diagnostics.CodeAnalysis;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;

using PowerPointLabs.ActionFramework.Common.Factory;
using PowerPointLabs.ActionFramework.Common.Log;
using PowerPointLabs.PictureSlidesLab.Views;
using PowerPointLabs.TextCollection;
using PowerPointLabs.Utils;

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
            ActionFramework.Common.Interface.ActionHandler actionHandler = ActionHandlerFactory.CreateInstance(control.Id, control.Tag);
            actionHandler.Execute(control.Id);
        }

        public bool GetEnabled(Office.IRibbonControl control)
        {
            if (IsAnyWindowOpen())
            {
                ActionFramework.Common.Interface.EnabledHandler enabledHandler = EnabledHandlerFactory.CreateInstance(control.Id, control.Tag);
                return enabledHandler.Get(control.Id);
            } 
            else 
            {
                return false;
            }
        }

        public string GetLabel(Office.IRibbonControl control)
        {
            ActionFramework.Common.Interface.LabelHandler labelHandler = LabelHandlerFactory.CreateInstance(control.Id, control.Tag);
            return labelHandler.Get(control.Id);
        }

        public string GetSupertip(Office.IRibbonControl control)
        {
            ActionFramework.Common.Interface.SupertipHandler supertipHandler = SupertipHandlerFactory.CreateInstance(control.Id, control.Tag);
            return supertipHandler.Get(control.Id);
        }

        public Bitmap GetImage(Office.IRibbonControl control)
        {
            ActionFramework.Common.Interface.ImageHandler imageHandler = ImageHandlerFactory.CreateInstance(control.Id, control.Tag);
            return imageHandler.Get(control.Id);
        }

        public string GetContent(Office.IRibbonControl control)
        {
            ActionFramework.Common.Interface.ContentHandler contentHandler = ContentHandlerFactory.CreateInstance(control.Id, control.Tag);
            return contentHandler.Get(control.Id);
        }

        public bool GetPressed(Office.IRibbonControl control)
        {
            ActionFramework.Common.Interface.PressedHandler pressedHandler = PressedHandlerFactory.CreateInstance(control.Id, control.Tag);
            return pressedHandler.Get(control.Id);
        }

        public void OnCheckBoxAction(Office.IRibbonControl control, bool pressed)
        {
            ActionFramework.Common.Interface.CheckBoxActionHandler checkBoxActionHandler = CheckBoxActionHandlerFactory.CreateInstance(control.Id, control.Tag);
            checkBoxActionHandler.Execute(control.Id, pressed);
        }

        #endregion

        #region Deprecated. Please only use Action Framework to support the feature.

#pragma warning disable 0618
        private Office.IRibbonUI _ribbon;
        // Initial bool value for whether the Drawing Tools Format Tab is disabled
        private bool DisableFormatTab { get; set; }
        // Initial bool value for whether images should be compressed
        private bool ShouldCompressImages { get; set; }

        #region IRibbonExtensibility Members

        public string GetCustomUI(string ribbonId)
        {
            return GetResourceText("PowerPointLabs.Ribbon1.xml");
        }

        #endregion

        #region Ribbon Callbacks
        //Create callback methods here. For more information about adding callback methods, select the Ribbon XML item in Solution Explorer and then press F1

        // Set the visibility of the Drawing Tools Format Tab
        public bool ToggleVisibleFormatTab(Office.IRibbonControl control)
        {
            return !DisableFormatTab;
        }

        // Toggles the boolean controlling the visibility of the Drawing Tools Format Tab
        public void ToggleVisibility(Office.IRibbonControl control, bool pressed)
        {
            DisableFormatTab = pressed;
            _ribbon.InvalidateControlMso("TabDrawingToolsFormat");
            _ribbon.InvalidateControl("VisibleFormatShapes");
        }

        // Toggles the boolean controlling whether images should be compressed.
        public void ToggleImageCompression(Office.IRibbonControl control, bool pressed)
        {
            ShouldCompressImages = pressed;
            _ribbon.InvalidateControl("ShouldCompressImagesCheckbox");
            GraphicsUtil.ShouldCompressPictureExport(ShouldCompressImages);
        }


        public void InitialiseVisibilityCheckbox()
        {
            _ribbon.InvalidateControl("VisibleFormatShapes");
        }
        
        public void InitialiseCompressImagesCheckbox()
        {
            _ribbon.InvalidateControl("ShouldCompressImagesCheckbox");
        }

        // Sets the default starting status of the checkbox (whether checked or not)
        public bool SetVisibility(Office.IRibbonControl control)
        {
            return DisableFormatTab;
        }

        // Sets the default starting status of the checkbox (whether checked or not)
        public bool SetImageCompression(Office.IRibbonControl control)
        {
            return ShouldCompressImages;
        }

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

            DisableFormatTab = new Boolean();
            ShouldCompressImages = new Boolean();

            _ribbon = ribbonUi;
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

        #region Button Labels
        public string GetPowerPointLabsAddInsTabLabel(Office.IRibbonControl control)
        {
            return CommonText.PowerPointLabsAddInsTabLabel;
        }

        public string GetCombineShapesLabel(Office.IRibbonControl control)
        {
            return CommonText.CombineShapesLabel;
        }

        public string GetPowerPointLabsMenuLabel(Office.IRibbonControl control)
        {
            return CommonText.PowerPointLabsMenuLabel;
        }
        # endregion

        public bool IsValidPresentation(PowerPoint.Presentation pres)
        {
            if (!Globals.ThisAddIn.VerifyVersion(pres))
            {
                MessageBox.Show(CommonText.ErrorVersionNotCompatible);
                return false;
            }

            return true;
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

        #region Feature: Combine Shapes
        public bool GetVisibilityForCombineShapes(Office.IRibbonControl control)
        {
            const string officeVersion2010 = "14.0";
            return Globals.ThisAddIn.Application.Version == officeVersion2010;
        }
        #endregion

        private static string GetResourceText(string resourceName)
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            string[] resourceNames = asm.GetManifestResourceNames();
            for (int i = 0; i < resourceNames.Length; i++)
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

        private bool IsAnyWindowOpen() 
        {
            return Globals.ThisAddIn.Application.Windows.Count > 0;
        }
        #endregion
    }
}
