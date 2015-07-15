using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using PowerPointLabs.DataSources;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using PowerPointLabs.Models;
using System.Drawing.Drawing2D;
using PPExtraEventHelper;
using Converters = PowerPointLabs.Converters;
using Microsoft.Office.Interop.PowerPoint;
using PowerPointLabs.ColorPicker;
using PowerPointLabs.Views;
using Microsoft.Office.Core;

namespace PowerPointLabs
{

    public partial class DrawingsPane : UserControl
    {
        public DrawingsPane()
        {
            InitializeComponent();

            BindDataToPanels();

            InitToolTipControl();
        }

        #region ToolTip
        private void InitToolTipControl()
        {
            //toolTip1.SetToolTip(panel1, TextCollection.ColorsLabText.MainColorBoxTooltips);
        }
        #endregion

        #region DataBindings
        private void BindDataToPanels()
        {
            //this.panel1.DataBindings.Add(new CustomBinding(
                //"BackColor",
                //dataSource,
                //"selectedColor",
                //new Converters.HSLColorToRGBColor()));
        }

        #endregion


        protected override CreateParams CreateParams
        {
            get
            {
                var createParams = base.CreateParams;
                createParams.ExStyle |= (int)Native.Message.WS_EX_COMPOSITED;  // Turn on WS_EX_COMPOSITED
                return createParams;
            }
        }
    }
}
