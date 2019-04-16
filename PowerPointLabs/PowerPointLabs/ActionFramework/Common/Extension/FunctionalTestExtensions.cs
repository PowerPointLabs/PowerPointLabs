using System;

using Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Tools;

using PowerPointLabs.Models;

namespace PowerPointLabs.ActionFramework.Common.Extension
{
    /// <summary>
    /// Provide Functional Test classes with powerpoint context.
    /// DO NOT USE THIS CLASS if you're not developing functional test,
    /// use ActionFramework instead.
    /// </summary>
    static class FunctionalTestExtensions
    {
#pragma warning disable 0618
        public static Application GetApplication()
        {
            return ActionFrameworkExtensions.GetApplication();
        }

        public static Presentations GetPresentations()
        {
            return ActionFrameworkExtensions.GetPresentations();
        }

        public static DocumentWindow GetCurrentWindow()
        {
            return ActionFrameworkExtensions.GetCurrentWindow();
        }

        public static Selection GetCurrentSelection()
        {
            return ActionFrameworkExtensions.GetCurrentSelection();
        }

        public static PowerPointSlide GetCurrentSlide()
        {
            return ActionFrameworkExtensions.GetCurrentSlide();
        }

        public static PowerPointPresentation GetCurrentPresentation()
        {
            return ActionFrameworkExtensions.GetCurrentPresentation();
        }

        public static Ribbon1 GetRibbonUi()
        {
            return ActionFrameworkExtensions.GetRibbonUi();
        }

        public static ThisAddIn GetAddIn()
        {
            return ActionFrameworkExtensions.GetAddIn();
        }

        /// <summary>
        /// Go to a slide
        /// </summary>
        /// <param name="slideIndex">1-based</param>
        public static void GotoSlide(int slideIndex)
        {
            ActionFrameworkExtensions.GotoSlide(slideIndex);
        }

        public static void ExecuteOfficeCommand(string commandMso)
        {
            ActionFrameworkExtensions.ExecuteOfficeCommand(commandMso);
        }

        public static void StartNewUndoEntry()
        {
            ActionFrameworkExtensions.StartNewUndoEntry();
        }

        public static CustomTaskPane GetTaskPane(Type taskPaneType)
        {
            return ActionFrameworkExtensions.GetTaskPane(taskPaneType);
        }

        public static CustomTaskPane RegisterTaskPane(Type taskPaneType, string taskPaneTitle, 
            EventHandler visibleChangeEventHandler = null, EventHandler dockPositionChangeEventHandler = null)
        {
            return ActionFrameworkExtensions.RegisterTaskPane(taskPaneType, taskPaneTitle,
                visibleChangeEventHandler, dockPositionChangeEventHandler);
        }
#pragma warning restore 0618
    }
}
