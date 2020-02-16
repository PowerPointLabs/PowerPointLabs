using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Controls;
using System.Windows.Media;

namespace PowerPointLabs.ColorThemes
{
    /// <summary>
    /// The ButtonTheme struct contains the colors that are used for the <see cref="Button"/> control.
    /// This includes the background, foreground and border colors when the Button is in its normal,
    /// mouse over, pressed and disabled states.
    /// </summary>
    public struct ButtonTheme
    {
        /// <summary>
        /// The background of a Button in its normal state.
        /// </summary>
        public Color NormalBackground { get; set; }
        /// <summary>
        /// The foreground of a Button in its normal state.
        /// </summary>
        public Color NormalForeground { get; set; }
        /// <summary>
        /// The border color of a Button in its normal state.
        /// </summary>
        public Color NormalBorderColor { get; set; }
        /// <summary>
        /// The background of a Button when the mouse is hovering over it.
        /// </summary>
        public Color MouseOverBackground { get; set; }
        /// <summary>
        /// The border color of a Button when the mouse is hovering over it.
        /// </summary>
        public Color MouseOverBorderColor { get; set; }
        /// <summary>
        /// The background of a pressed Button.
        /// </summary>
        public Color PressedBackground { get; set; }
        /// <summary>
        /// The border color of a pressed Button.
        /// </summary>
        public Color PressedBorderColor { get; set; }
        /// <summary>
        /// The background of a disabled Button.
        /// </summary>
        public Color DisabledBackground { get; set; }
        /// <summary>
        /// The foreground of a disabled Button.
        /// </summary>
        public Color DisabledForeground { get; set; }
        /// <summary>
        /// The border color of a disabled button.
        /// </summary>
        public Color DisabledBorderColor { get; set; }
    }
}
