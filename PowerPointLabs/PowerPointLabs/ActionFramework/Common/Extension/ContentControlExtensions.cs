using System;
using System.Windows.Controls;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Tools;
using PowerPointLabs.Models;

namespace PowerPointLabs.ActionFramework.Common.Extension
{
    /// <summary>
    /// Extend ContentControl with powerpoint context
    /// </summary>
    static class ContentControlExtensions
    {
        public static DocumentWindow GetCurrentWindow(this ContentControl handler)
        {
            return ActionFrameworkExtensions.GetCurrentWindow();
        }

        public static Selection GetCurrentSelection(this ContentControl handler)
        {
            return ActionFrameworkExtensions.GetCurrentSelection();
        }

        public static PowerPointSlide GetCurrentSlide(this ContentControl handler)
        {
            return ActionFrameworkExtensions.GetCurrentSlide();
        }

        public static PowerPointPresentation GetCurrentPresentation(this ContentControl handler)
        {
            return ActionFrameworkExtensions.GetCurrentPresentation();
        }

        public static Ribbon1 GetRibbonUi(this ContentControl handler)
        {
            return ActionFrameworkExtensions.GetRibbonUi();
        }

        /// <summary>
        /// Go to a slide
        /// </summary>
        /// <param name="handler"></param>
        /// <param name="slideIndex">1-based</param>
        public static void GotoSlide(this ContentControl handler, int slideIndex)
        {
            ActionFrameworkExtensions.GotoSlide(slideIndex);
        }

        public static void ExecuteOfficeCommand(this ContentControl handler, string commandMso)
        {
            ActionFrameworkExtensions.ExecuteOfficeCommand(commandMso);
        }

        public static void StartNewUndoEntry(this ContentControl handler)
        {
            ActionFrameworkExtensions.StartNewUndoEntry();
        }

        public static CustomTaskPane GetTaskPane(this ContentControl handler, Type taskPaneType)
        {
            return ActionFrameworkExtensions.GetTaskPane(taskPaneType);
        }

        public static CustomTaskPane RegisterTaskPane(this ContentControl handler, Type taskPaneType, string taskPaneTitle, 
            EventHandler visibleChangeEventHandler = null, EventHandler dockPositionChangeEventHandler = null)
        {
            return ActionFrameworkExtensions.RegisterTaskPane(taskPaneType, taskPaneTitle,
                visibleChangeEventHandler, dockPositionChangeEventHandler);
        }
    }
}
