using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using PowerPointLabs.Models;
using PPExtraEventHelper;
using Shape = Microsoft.Office.Interop.PowerPoint.Shape;
using Graphics = PowerPointLabs.Utils.Graphics;

namespace PowerPointLabs.DrawingsLab
{
    internal class DrawingsLabMain
    {
        public static void SwitchToLineTool()
        {
            // This should trigger the line tool.
            // see https://github.com/PowerPointLabs/powerpointlabs/blob/master/PowerPointLabs/PowerPointLabs/ThisAddIn.cs#L1381
            //TODO: Placeholder code. This just triggers the property window.
            Native.SendMessage(
                Process.GetCurrentProcess().MainWindowHandle,
                (uint)Native.Message.WM_COMMAND,
                new IntPtr(0x8F),
                IntPtr.Zero
                );
        }

        public static void HideTool()
        {
            var selection = PowerPointCurrentPresentationInfo.CurrentSelection;
            if (selection.Type != PpSelectionType.ppSelectionShapes) return;

            foreach (var shape in selection.ShapeRange.Cast<Shape>())
            {
                shape.Visible = MsoTriState.msoFalse;
            }
        }

        public static void CloneTool()
        {
            var selection = PowerPointCurrentPresentationInfo.CurrentSelection;
            if (selection.Type != PpSelectionType.ppSelectionShapes) return;

            PowerPointCurrentPresentationInfo.CurrentSlide.CopyShapesToSlide(selection.ShapeRange);
        }

        public static void MultiCloneTool()
        {
            var selection = PowerPointCurrentPresentationInfo.CurrentSelection;
            if (selection.Type != PpSelectionType.ppSelectionShapes) return;
            var shapeList = selection.ShapeRange.Cast<Shape>().ToList();

            if (shapeList.Count % 2 != 0)
            {
                Error("There must be two sets of shapes selected.");
                return;
            }

            int clones = ShowNumericDialog("Number of copies:", "Multi-Clone") - 1;
            if (clones <= 0) return;

            int midpoint = shapeList.Count / 2;
            for (int i = 0; i < shapeList.Count / 2; ++i)
            {
                // Do the cloning for every pair of shapes (i, midpoint+i)
                var firstShape = shapeList[i];
                var secondShape = shapeList[midpoint + i];

                for (int j = 0; j < clones; ++j)
                {
                    var newShape = firstShape.Duplicate()[1];
                    int index = j + 1;

                    newShape.Left = secondShape.Left + (secondShape.Left - firstShape.Left) * index;
                    newShape.Top = secondShape.Top + (secondShape.Top - firstShape.Top) * index;
                    newShape.Rotation = secondShape.Rotation + (secondShape.Rotation - firstShape.Rotation) * index;
                }
            }
        }

        private static void Error(string message)
        {
            // for now do nothing.
        }

        private static int ShowNumericDialog(string text, string caption)
        {
            var prompt = new Form()
            {
                FormBorderStyle = FormBorderStyle.FixedDialog,
                MinimizeBox = false,
                MaximizeBox = false,
                Width = 160,
                Height = 130,
                Text = caption,
                StartPosition = FormStartPosition.CenterScreen,
            };

            var cancel = new Button();
            cancel.Click += (sender, e) => prompt.Close();
            prompt.CancelButton = cancel;

            var textLabel = new Label()
            {
                Top = 10,
                Text = text,
                TextAlign = ContentAlignment.MiddleCenter,
                AutoSize = false,
                Width = prompt.Width
            };

            var textBox = new NumericUpDown() { Left = 20, Top = 40, Width = 120, Height = 80, Text = "5" };
            var confirmation = new Button() { Text = "Ok", Left = 30, Top = 70, Width = 100, DialogResult = DialogResult.OK };
            confirmation.Click += (sender, e) => { prompt.Close(); };

            prompt.Controls.Add(textBox);
            prompt.Controls.Add(confirmation);
            prompt.Controls.Add(textLabel);
            prompt.AcceptButton = confirmation;

            textBox.Select(0, textBox.Text.Length);

            if (prompt.ShowDialog() == DialogResult.OK)
            {
                int inputValue;
                if (int.TryParse(textBox.Text, out inputValue))
                {
                    return inputValue;
                }
            }
            return -1;
        }
    }
}
