using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace PowerPointLabs.ColorThemes
{
    /// <summary>
    /// The BaseStylesDictionary class is a ResourceDictionary that contains all of the base
    /// styles of Controls.
    /// </summary>
    /// <remarks>
    /// The main purpose of this class is to be able to design styles that are based off of
    /// the base styles defined in the Resources/Themes folder. This allows them to maintain
    /// the colour changing properties upon the ColorTheme being changed.
    /// 
    /// That said, the main reason of creating this helper class is so that the path to the
    /// Resources/Themes folder doesn't have to be written out every time.
    /// 
    /// <code>
    /// <UserControl 
    ///     x:Class="..."
    ///     xmlns="..."
    ///     xmlns:theme="clr-namespace:PowerPointLabs.ColorThemes">
    ///     <UserControl.Resources>
    ///         <ResourceDictionary>
    ///             <ResourceDictionary.MergedDictionaries>
    ///                 <theme:TemplateStyleDictionary/>
    ///             </ResourceDictionary.MergedDictionaries>    
    ///         </ResourceDictionary>
    ///     </UserControl.Resources>
    ///     <Button>
    ///         <Button.Style>
    ///             <Style TargetType="{x:Type Button}" BasedOn="{StaticResource BaseButtonStyle}">
    ///                 <Style.Triggers>
    ///                     <Trigger Property="IsEnabled" Value="False">
    ///                         <Setter Property="Opacity" Value="0.3"/>
    ///                     </Trigger>
    ///                 </Style.Triggers>
    ///             </Style>
    ///         </Button.Style>
    ///     </Button>
    /// </UserControl>
    /// </code>
    /// </remarks>
    public class BaseStylesDictionary : ResourceDictionary
    {
        public static readonly string PathToThemesFolder = "pack://application:,,,/PowerPointLabs;component/Resources/Themes/";

        public BaseStylesDictionary()
        {
            Source = new Uri(PathToThemesFolder + "TemplateStyles.xaml", UriKind.RelativeOrAbsolute);
        }
    }
}
