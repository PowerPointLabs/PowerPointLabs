using System;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Tools;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.Models;

namespace PowerPointLabs.ActionFramework.Common.Extension
{
    /// <summary>
    /// Extend action handler with powerpoint context.
    /// Type `this` in the class/subclass of ActionHandler to access the APIs below.
    /// </summary>
    [Obsolete("DO NOT use this class in your feature! Used only by Action Framework.")]
    static class ActionHandlerExtensions
    {
#pragma warning disable 0618
        public static Application GetApplication(this ActionHandler handler)
        {
            return ActionFrameworkExtensions.GetApplication();
        }

        public static Presentations GetPresentations(this ActionHandler handler)
        {
            return ActionFrameworkExtensions.GetPresentations();
        }

        public static DocumentWindow GetCurrentWindow(this ActionHandler handler)
        {
            return ActionFrameworkExtensions.GetCurrentWindow();
        }

        public static Selection GetCurrentSelection(this ActionHandler handler)
        {
            return ActionFrameworkExtensions.GetCurrentSelection();
        }

        public static PowerPointSlide GetCurrentSlide(this ActionHandler handler)
        {
            return ActionFrameworkExtensions.GetCurrentSlide();
        }

        public static PowerPointPresentation GetCurrentPresentation(this ActionHandler handler)
        {
            return ActionFrameworkExtensions.GetCurrentPresentation();
        }

        public static Ribbon1 GetRibbonUi(this ActionHandler handler)
        {
            return ActionFrameworkExtensions.GetRibbonUi();
        }

        /// <summary>
        /// Go to a slide
        /// </summary>
        /// <param name="handler"></param>
        /// <param name="slideIndex">1-based</param>
        public static void GotoSlide(this ActionHandler handler, int slideIndex)
        {
            ActionFrameworkExtensions.GotoSlide(slideIndex);
        }

        public static void ExecuteOfficeCommand(this ActionHandler handler, string commandMso)
        {
            ActionFrameworkExtensions.ExecuteOfficeCommand(commandMso);
        }

        public static void StartNewUndoEntry(this ActionHandler handler)
        {
            ActionFrameworkExtensions.StartNewUndoEntry();
        }

        public static CustomTaskPane GetTaskPane(this ActionHandler handler, Type taskPaneType)
        {
            return ActionFrameworkExtensions.GetTaskPane(taskPaneType);
        }

        public static CustomTaskPane RegisterTaskPane(this ActionHandler handler, Type taskPaneType, string taskPaneTitle, 
            EventHandler visibleChangeEventHandler = null, EventHandler dockPositionChangeEventHandler = null)
        {
            return ActionFrameworkExtensions.RegisterTaskPane(taskPaneType, taskPaneTitle,
                visibleChangeEventHandler, dockPositionChangeEventHandler);
        }
#pragma warning restore 0618
    }
}
