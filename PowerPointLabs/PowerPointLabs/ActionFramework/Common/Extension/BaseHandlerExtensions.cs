using System;

using Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Tools;

using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.Models;

namespace PowerPointLabs.ActionFramework.Common.Extension
{
    /// <summary>
    /// Extend base handler with powerpoint context.
    /// Type `this` in the class/subclass of BaseHandler to access the APIs below.
    /// </summary>
    [Obsolete("DO NOT use this class in your feature! Used only by Action Framework.")]
    static class BaseHandlerExtensions
    {
#pragma warning disable 0618
        public static Application GetApplication(this BaseHandler handler)
        {
            return ActionFrameworkExtensions.GetApplication();
        }

        public static Presentations GetPresentations(this BaseHandler handler)
        {
            return ActionFrameworkExtensions.GetPresentations();
        }

        public static DocumentWindow GetCurrentWindow(this BaseHandler handler)
        {
            return ActionFrameworkExtensions.GetCurrentWindow();
        }

        public static Selection GetCurrentSelection(this BaseHandler handler)
        {
            return ActionFrameworkExtensions.GetCurrentSelection();
        }

        public static PowerPointSlide GetCurrentSlide(this BaseHandler handler)
        {
            return ActionFrameworkExtensions.GetCurrentSlide();
        }

        public static PowerPointPresentation GetCurrentPresentation(this BaseHandler handler)
        {
            return ActionFrameworkExtensions.GetCurrentPresentation();
        }

        public static Ribbon1 GetRibbonUi(this BaseHandler handler)
        {
            return ActionFrameworkExtensions.GetRibbonUi();
        }

        public static ThisAddIn GetAddIn(this BaseHandler handler)
        {
            return ActionFrameworkExtensions.GetAddIn();
        }

        /// <summary>
        /// Go to a slide
        /// </summary>
        /// <param name="handler"></param>
        /// <param name="slideIndex">1-based</param>
        public static void GotoSlide(this BaseHandler handler, int slideIndex)
        {
            ActionFrameworkExtensions.GotoSlide(slideIndex);
        }

        public static void ExecuteOfficeCommand(this BaseHandler handler, string commandMso)
        {
            ActionFrameworkExtensions.ExecuteOfficeCommand(commandMso);
        }

        public static void StartNewUndoEntry(this BaseHandler handler)
        {
            ActionFrameworkExtensions.StartNewUndoEntry();
        }

        public static CustomTaskPane GetTaskPane(this BaseHandler handler, Type taskPaneType)
        {
            return ActionFrameworkExtensions.GetTaskPane(taskPaneType);
        }

        public static CustomTaskPane RegisterTaskPane(this BaseHandler handler, Type taskPaneType, string taskPaneTitle, 
            EventHandler visibleChangeEventHandler = null, EventHandler dockPositionChangeEventHandler = null)
        {
            return ActionFrameworkExtensions.RegisterTaskPane(taskPaneType, taskPaneTitle,
                visibleChangeEventHandler, dockPositionChangeEventHandler);
        }
#pragma warning restore 0618
    }
}
