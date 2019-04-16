using System;
using System.Windows.Forms;

using Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Tools;

using PowerPointLabs.Models;

using Application = Microsoft.Office.Interop.PowerPoint.Application;

namespace PowerPointLabs.ActionFramework.Common.Extension
{
    /// <summary>
    /// Provide powerpoint context for Action Framework
    /// </summary>
    [Obsolete("DO NOT use this class in your feature! Used only by Action Framework.")]
    class ActionFrameworkExtensions
    {
#pragma warning disable 0618
        public static Application GetApplication()
        {
            return Globals.ThisAddIn.Application;
        }

        public static Presentations GetPresentations()
        {
            return Globals.ThisAddIn.Application.Presentations;
        }

        public static DocumentWindow GetCurrentWindow()
        {
            return Globals.ThisAddIn.Application.ActiveWindow;
        }

        /// <summary>
        /// Gets current selection in Active Window. Returns null if selection cannot be retrieved.
        /// </summary>
        public static Selection GetCurrentSelection()
        {
            return PowerPointCurrentPresentationInfo.CurrentSelection;
        }

        public static PowerPointSlide GetCurrentSlide()
        {
            return PowerPointCurrentPresentationInfo.CurrentSlide;
        }

        public static PowerPointPresentation GetCurrentPresentation()
        {
            return PowerPointPresentation.Current;
        }

        public static Ribbon1 GetRibbonUi()
        {
            return Globals.ThisAddIn.Ribbon;
        }

        public static ThisAddIn GetAddIn()
        {
            return Globals.ThisAddIn;
        }

        /// <summary>
        /// Go to a slide
        /// </summary>
        /// <param name="slideIndex">1-based</param>
        public static void GotoSlide(int slideIndex)
        {
            Globals.ThisAddIn.Application.ActiveWindow.View.GotoSlide(slideIndex);
        }

        public static void ExecuteOfficeCommand(string commandMso)
        {
            Microsoft.Office.Core.CommandBars commandBars = Globals.ThisAddIn.Application.CommandBars;
            commandBars.ExecuteMso(commandMso);
        }

        public static void StartNewUndoEntry()
        {
            Globals.ThisAddIn.Application.StartNewUndoEntry();
        }

        public static CustomTaskPane GetTaskPane(Type taskPaneType)
        {
            return Globals.ThisAddIn.GetActivePane(taskPaneType);
        }

        public static CustomTaskPane RegisterTaskPane(Type taskPaneType, string taskPaneTitle, 
            EventHandler visibleChangeEventHandler = null, EventHandler dockPositionChangeEventHandler = null)
        {
            try
            {
                CustomTaskPane taskPane = Globals.ThisAddIn.GetActivePane(taskPaneType);
                if (taskPane != null)
                {
                    return taskPane;
                }

                UserControl taskPaneControl = (UserControl) Activator.CreateInstance(taskPaneType);
                if (taskPaneControl == null)
                {
                    throw new InvalidCastException("Failed to convert " + taskPaneType + " to UserControl.");
                }

                DocumentWindow activeWindow = Globals.ThisAddIn.Application.ActiveWindow;

                return Globals.ThisAddIn.RegisterTaskPane(taskPaneControl, taskPaneTitle, activeWindow,
                    visibleChangeEventHandler, dockPositionChangeEventHandler);
            }
            catch (Exception e)
            {
                Log.Logger.LogException(e, "RegisterTaskPane_Extension");
                Views.ErrorDialogBox.ShowDialog("PowerPointLabs", e.Message, e);
                return null;
            }
        }
#pragma warning restore 0618
    }
}
