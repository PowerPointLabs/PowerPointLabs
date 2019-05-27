using System.Drawing;
using System.Windows.Controls;

using Microsoft.Office.Core;

using PowerPointLabs.Utils;

namespace PowerPointLabs.TooltipsLab.Views
{
    /// <summary>
    /// Interaction logic for TooltipsLabSettingsShapeEntry.xaml
    /// </summary>
    public partial class TooltipsLabSettingsShapeEntry : UserControl
    {
        private MsoAutoShapeType type;

        #region Constructors

        public TooltipsLabSettingsShapeEntry(MsoAutoShapeType type, Bitmap image)
        {
            InitializeComponent();
            Type = type;
            imageBox.Source = CommonUtil.CreateBitmapSource(image);
        }

        #endregion

        #region Properties

        public MsoAutoShapeType Type
        {
            get
            {
                return type;
            }
            set
            {
                type = value;
                string nameForDisplay = value.ToString().Replace(
                    TooltipsLabConstants.ShapeNameHeader, "");
                textBlock.Text = nameForDisplay;
                ToolTip = nameForDisplay;
            }
        }

        #endregion
    }
}