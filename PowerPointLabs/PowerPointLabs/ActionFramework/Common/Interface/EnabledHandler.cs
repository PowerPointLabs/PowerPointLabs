using System;

using Microsoft.Office.Interop.PowerPoint;

using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.Utils;

namespace PowerPointLabs.ActionFramework.Common.Interface
{
    /// <summary>
    /// Handler that handles GetEnabled class
    /// </summary>
    public abstract class EnabledHandler : BaseHandler
    {
        public bool Get(string ribbonId)
        {
            try
            {
                return GetEnabled(ribbonId);
            }
            catch (Exception e)
            {
                Log.Logger.LogException(e, ribbonId);
                Views.ErrorDialogWrapper.ShowDialog("PowerPointLabs", e.Message, e);
                return false;
            }
        }

        protected abstract bool GetEnabled(string ribbonId);
        
        protected bool HasPlaceholderInSelection()
        {
            Selection selection = this.GetCurrentSelection();
            
            if (selection.Type != PpSelectionType.ppSelectionShapes)
            {
                return false;
            }

            foreach (Shape shape in selection.ShapeRange)
            {
                if (shape.Type == Microsoft.Office.Core.MsoShapeType.msoPlaceholder)
                {
                    return true;
                }
            }

            return false;
        }

        protected bool IsSelectionSingleShape()
        {
            Selection selection = this.GetCurrentSelection();

            if (selection.HasChildShapeRange)
            {
                return selection.ChildShapeRange.Count == 1 &&
                    selection.ChildShapeRange.Type == Microsoft.Office.Core.MsoShapeType.msoAutoShape;
            }

            return selection.Type == PpSelectionType.ppSelectionShapes && 
                selection.ShapeRange.Count == 1;
        }

        protected bool IsSelectionMultipleOrGroup()
        {
            Selection selection = this.GetCurrentSelection();

            if (selection.Type != PpSelectionType.ppSelectionShapes)
            {
                return false;
            }

            if (selection.ShapeRange.Count > 1)
            {
                return true;
            }

            if (Graphics.IsAGroup(selection.ShapeRange[1]))
            {
                return true;
            }

            return false;
        }

        protected bool IsSelectionMultipleSameShapeType()
        {
            bool isEnabled = false;
            var selection = this.GetCurrentSelection();
            if (selection.Type == PpSelectionType.ppSelectionNone ||
                selection.Type == PpSelectionType.ppSelectionSlides)
            {
                return isEnabled;
            }

            var shape = selection.ShapeRange[1];

            if (selection.ShapeRange.Count > 1)
            {
                foreach (Shape tempShape in selection.ShapeRange)
                {
                    if (shape.Type == tempShape.Type)
                    {
                        isEnabled = true;
                        //Ribbon.ZoomButtonEnabled = true;
                    }
                    if (shape.Type == Microsoft.Office.Core.MsoShapeType.msoAutoShape &&
                        shape.AutoShapeType != tempShape.AutoShapeType)
                    {
                        isEnabled = false;
                        //Ribbon.ZoomButtonEnabled = false;
                        break;
                    }
                }
            }
            return isEnabled;
        }

        protected bool IsSelectionChildShapeRange()
        {
            Selection selection = this.GetCurrentSelection();
            return selection.HasChildShapeRange;
        }
    }
}
