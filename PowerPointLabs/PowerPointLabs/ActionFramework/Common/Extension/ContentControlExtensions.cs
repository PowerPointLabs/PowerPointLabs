using System;
using System.Windows.Controls;

using Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Tools;

using PowerPointLabs.Models;

namespace PowerPointLabs.ActionFramework.Common.Extension
{
    /// <summary>
    /// Extend ContentControl with powerpoint context.
    /// Type `this` in the class/subclass of ContentControl (WPF UserControl) to access the APIs below.
    /// </summary>
    [Obsolete("DO NOT use this class in your feature! Used only by Action Framework.")]
    static class ContentControlExtensions
    {
#pragma warning disable 0618
        public static Application GetApplication(this ContentControl control)
        {
            return ActionFrameworkExtensions.GetApplication();
        }

        public static Presentations GetPresentations(this ContentControl control)
        {
            return ActionFrameworkExtensions.GetPresentations();
        }

        public static DocumentWindow GetCurrentWindow(this ContentControl control)
        {
            return ActionFrameworkExtensions.GetCurrentWindow();
        }

        public static Selection GetCurrentSelection(this ContentControl control)
        {
            return ActionFrameworkExtensions.GetCurrentSelection();
        }

        public static PowerPointSlide GetCurrentSlide(this ContentControl control)
        {
            return ActionFrameworkExtensions.GetCurrentSlide();
        }

        public static PowerPointPresentation GetCurrentPresentation(this ContentControl control)
        {
            return ActionFrameworkExtensions.GetCurrentPresentation();
        }

        public static Ribbon1 GetRibbonUi(this ContentControl control)
        {
            return ActionFrameworkExtensions.GetRibbonUi();
        }

        public static ThisAddIn GetAddIn(this ContentControl control)
        {
            return ActionFrameworkExtensions.GetAddIn();
        }

        /// <summary>
        /// Go to a slide
        /// </summary>
        /// <param name="control"></param>
        /// <param name="slideIndex">1-based</param>
        public static void GotoSlide(this ContentControl control, int slideIndex)
        {
            ActionFrameworkExtensions.GotoSlide(slideIndex);
        }

        public static void ExecuteOfficeCommand(this ContentControl control, string commandMso)
        {
            ActionFrameworkExtensions.ExecuteOfficeCommand(commandMso);
        }

        public static void StartNewUndoEntry(this ContentControl control)
        {
            ActionFrameworkExtensions.StartNewUndoEntry();
        }

        public static CustomTaskPane GetTaskPane(this ContentControl control, Type taskPaneType)
        {
            return ActionFrameworkExtensions.GetTaskPane(taskPaneType);
        }

        public static CustomTaskPane RegisterTaskPane(this ContentControl control, Type taskPaneType, string taskPaneTitle, 
            EventHandler visibleChangeEventcontrol = null, EventHandler dockPositionChangeEventcontrol = null)
        {
            return ActionFrameworkExtensions.RegisterTaskPane(taskPaneType, taskPaneTitle,
                visibleChangeEventcontrol, dockPositionChangeEventcontrol);
        }
#pragma warning restore 0618
    }
}
