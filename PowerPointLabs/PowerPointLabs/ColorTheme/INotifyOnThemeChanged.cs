using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerPointLabs.ColorThemes
{
    /// <summary>
    /// The INotifyOnThemeChanged interface is to indicate that the class implementing this interface
    /// should be notified when the ColorTheme has changed.
    /// </summary>
    /// <remarks>
    /// This interface should only be implemented by classes that are subscribed to the 
    /// <see cref="ThemeManager.ColorThemeChanged"/> event via the
    /// <see cref="Extensions.ThemeExtensions.ApplyTheme(System.Windows.FrameworkElement, object, ColorTheme)"/>
    /// method. The <see cref="OnThemeChanged(ColorTheme)"/> method only gets called via that ApplyTheme method.
    /// </remarks>
    public interface INotifyOnThemeChanged
    {
        /// <summary>
        /// Notifies the class that the Color Theme has changed.
        /// </summary>
        /// <param name="updatedColorTheme">The updated ColorTheme</param>
        void OnThemeChanged(ColorTheme updatedColorTheme);
    }
}
