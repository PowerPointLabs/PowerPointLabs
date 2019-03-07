using System.Windows.Controls;

using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using PowerPointLabs.Utils;

using Shape = Microsoft.Office.Interop.PowerPoint.Shape;

namespace PowerPointLabs.TooltipsLab
{
    class TextBoxWithDefaultText
    {
        private const string DefaultText = "Enter Text Here";

        private bool isTextEdited = false;
        private Shape textBox;

        private TextBoxWithDefaultText()
        {

        }

        public static Shape CreateTextBox(Slide nativeSlide, MsoTextOrientation orientation, float left, float top, float width, float height)
        {
            return new TextBoxWithDefaultText().AddTextBox(nativeSlide, orientation, left, top, width, height);
        }

        public Shape AddTextBox(Slide nativeSlide, MsoTextOrientation orientation, float left, float top, float width, float height)
        {
            textBox = nativeSlide.Shapes.AddTextbox(orientation, left, top, width, height);
            nativeSlide.Application.WindowSelectionChange += SelectionChanged;
            return textBox;
        }

        private void SelectionChanged(Selection sel)
        {
            if (IsTextBoxSelected(sel))
            {
                HandleSelection();
            }
            else
            {
                HandleDeselection();
            }
        }

        private bool IsTextBoxSelected(Selection sel)
        {
            if (!(sel.Type == PpSelectionType.ppSelectionShapes) || !ShapeUtil.IsSelectionSingleShape(sel))
            {
                return false;
            }
            if (sel.ShapeRange[1] == textBox)
            {
                return true;
            }
            Shape shapeGroup = sel.ShapeRange[1];
            if (shapeGroup.Type != MsoShapeType.msoGroup)
            {
                return false;
            }
            foreach (Shape shape in shapeGroup.GroupItems)
            {
                if (shape == textBox)
                {
                    return true;
                }
            }
            return false;
        }

        private void HandleSelection()
        {
            TextRange textRange = textBox.TextFrame.TextRange;
            if (textRange.Text == DefaultText && !isTextEdited)
            {
                textRange.Text = "";
            }
        }

        private void HandleDeselection()
        {
            TextRange textRange = textBox.TextFrame.TextRange;
            if (textRange.Text == "")
            {
                isTextEdited = false;
                textRange.Text = DefaultText;
            }
        }
    }
}
