using System;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Tools;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.Models;

namespace PowerPointLabs.ActionFramework.Common.Extension
{
    /// <summary>
    /// Extend action handler with powerpoint context.
    /// Type `this` in the class/subclass of CheckBoxActionHandler to access the APIs below.
    /// </summary>
    [Obsolete("DO NOT use this class in your feature! Used only by Action Framework.")]
    static class CheckBoxActionHandlerExtensions
    {
#pragma warning disable 0618
        public static Application GetApplication(this CheckBoxActionHandler handler)
        {
            return ActionFrameworkExtensions.GetApplication();
        }

        public static Presentations GetPresentations(this CheckBoxActionHandler handler)
        {
            return ActionFrameworkExtensions.GetPresentations();
        }

        public static DocumentWindow GetCurrentWindow(this CheckBoxActionHandler handler)
        {
            return ActionFrameworkExtensions.GetCurrentWindow();
        }

        public static Selection GetCurrentSelection(this CheckBoxActionHandler handler)
        {
            return ActionFrameworkExtensions.GetCurrentSelection();
        }

        public static PowerPointSlide GetCurrentSlide(this CheckBoxActionHandler handler)
        {
            return ActionFrameworkExtensions.GetCurrentSlide();
        }

        public static PowerPointPresentation GetCurrentPresentation(this CheckBoxActionHandler handler)
        {
            return ActionFrameworkExtensions.GetCurrentPresentation();
        }

        public static Ribbon1 GetRibbonUi(this CheckBoxActionHandler handler)
        {
            return ActionFrameworkExtensions.GetRibbonUi();
        }

        /// <summary>
        /// Go to a slide
        /// </summary>
        /// <param name="handler"></param>
        /// <param name="slideIndex">1-based</param>
        public static void GotoSlide(this CheckBoxActionHandler handler, int slideIndex)
        {
            ActionFrameworkExtensions.GotoSlide(slideIndex);
        }

        public static void ExecuteOfficeCommand(this CheckBoxActionHandler handler, string commandMso)
        {
            ActionFrameworkExtensions.ExecuteOfficeCommand(commandMso);
        }

        public static void StartNewUndoEntry(this CheckBoxActionHandler handler)
        {
            ActionFrameworkExtensions.StartNewUndoEntry();
        }

        public static CustomTaskPane GetTaskPane(this CheckBoxActionHandler handler, Type taskPaneType)
        {
            return ActionFrameworkExtensions.GetTaskPane(taskPaneType);
        }

        public static CustomTaskPane RegisterTaskPane(this CheckBoxActionHandler handler, Type taskPaneType, string taskPaneTitle, 
            EventHandler visibleChangeEventHandler = null, EventHandler dockPositionChangeEventHandler = null)
        {
            return ActionFrameworkExtensions.RegisterTaskPane(taskPaneType, taskPaneTitle,
                visibleChangeEventHandler, dockPositionChangeEventHandler);
        }
#pragma warning restore 0618
    }
}
