﻿using System.Windows;

using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.Action.PasteLab
{
    [ExportActionRibbonId("pasteAndReplace")]
    class PasteLabReplaceActionHandler : ActionHandler
    {
        protected override void ExecuteAction(string ribbonId)
        {
            var slide = this.GetCurrentSlide();
            var selection = this.GetCurrentSelection();
            bool clipboardIsEmpty = (Clipboard.GetDataObject() == null);

            PowerPointLabs.PasteLab.PasteLabMain.PasteAndReplace(slide, clipboardIsEmpty, selection);
        }
    }
}