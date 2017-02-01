using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using Microsoft.Office.Interop.PowerPoint;
using PPExtraEventHelper;
using PowerPointLabs.SyncLab;
using PowerPointLabs.SyncLab.ObjectFormats;
using System.Drawing;
using System.Reflection;
using System.Diagnostics;

namespace PowerPointLabs
{
    public partial class SyncLabPane : UserControl
    {
#pragma warning disable 0618

        private bool _firstTimeLoading = true;

        private static readonly List<Type> OBJECT_FORMATS = InitializeObjectFormats();

        # region Constructors
        public SyncLabPane()
        {
            SetStyle(ControlStyles.UserPaint | ControlStyles.DoubleBuffer | ControlStyles.AllPaintingInWmPaint, true);
            InitializeComponent();
        }

        private static List<Type> InitializeObjectFormats()
        {
            List<Type> objectFormats = new List<Type>();
            // Only add ObjectFormat types
            //objectFormats.Add(typeof(SyncLab.ObjectFormats.LineFormat));
            objectFormats.Add(typeof(SyncLab.ObjectFormats.FillFormat));
            //objectFormats.Add(typeof(SyncLab.ObjectFormats.FillFormat));
            foreach (Type t in objectFormats) // Ensure all types in list are ObjectFormat types
            {
                Debug.Assert(t.IsSubclassOf(typeof(ObjectFormat)), "Not all types are subclasses of the ObjectFormat class");
            }
            return objectFormats;
        }
        # endregion

        # region API
        public void PaneReload(bool forceReload = false)
        {
            if (!_firstTimeLoading && !forceReload)
            {
                return;
            }

            _firstTimeLoading = false;
        }

        public void CopyFormat()
        {
            ShapeRange selectedShapes = Globals.ThisAddIn.Application.ActiveWindow.Selection.ShapeRange;
            CopyFormat(selectedShapes);
        }

        public void CopyFormat(ShapeRange shapes)
        {
            foreach (Shape shape in shapes)
            {
                CopyFormat(shape);
            }
        }

        public void CopyFormat(Shape shape)
        {
            // Clear checked items
            for (int i = 0; i < syncLabListBox.Items.Count; i++)
            {
                syncLabListBox.Items[i].Checked = false;
            }
            // Add all the formats
            foreach (Type formatClass in OBJECT_FORMATS)
            {
                ConstructorInfo cInfo = formatClass.GetConstructor(new[] { typeof(Shape) });
                object newObject = cInfo.Invoke(new object[] { shape });
                Debug.Assert(newObject is ObjectFormat, "Object instantiated is not an ObjectFormat");
                ObjectFormat newFormat = (ObjectFormat)newObject;
                syncLabListBox.AddFormat(newFormat);
                syncLabListBox.Items[0].Checked = true;
            }
        }

        public void PasteFormat()
        {
            ShapeRange selectedShapes = Globals.ThisAddIn.Application.ActiveWindow.Selection.ShapeRange;
            PasteFormat(selectedShapes);
        }

        public void PasteFormat(ShapeRange shapes)
        {
            foreach (Shape shape in shapes)
            {
                PasteFormat(shape);
            }
        }

        public void PasteFormat(Shape shape)
        {
            List<int> checkedIndices = syncLabListBox.CheckedIndices.Cast<int>().ToList<int>();
            checkedIndices.Sort();
            checkedIndices.Reverse();
            for (int i = 0; i < checkedIndices.Count; i++)
            {
                ObjectFormat format = syncLabListBox.GetFormat(checkedIndices[i]);
                format.ApplyTo(shape);
            }
        }

        #endregion

        #region Functional Test APIs

        public void AddStyleToList(ObjectFormat format)
        {
            syncLabListBox.AddFormat(format);
        }

        #endregion

        #region GUI Handlers
        private void CopyButton_Click(object sender, EventArgs e)
        {
            CopyFormat();
        }

        private void PasteButton_Click(object sender, EventArgs e)
        {
            PasteFormat();
        }
        #endregion
    }
}
