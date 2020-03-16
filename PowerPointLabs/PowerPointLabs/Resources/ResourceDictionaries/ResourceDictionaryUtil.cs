using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace PowerPointLabs.Resources.ResourceDictionaries
{
    public static class ResourceDictionaryUtil
    {
        /// <summary>
        /// The path to the ResourceDictionaries folder when used in a Uri.
        /// </summary>
        public static readonly string PathToResourceDictionaries = "PowerPointLabs;component/Resources/ResourceDictionaries/";

        /// <summary>
        /// Retrieves the resource given by the specified key in the Resource Dictionary with the 
        /// specified name. The Resource Dictionary must be in the Resources/ResourceDictionaries
        /// directory.
        /// </summary>
        /// <remarks>
        /// If the specified key does not exist in the specified Resource Dictionary, this method
        /// will return null.
        /// </remarks>
        /// <param name="resourceDictionaryName">The name of the resource dictionary.</param>
        /// <param name="key">The key of the resource to retrieve.</param>
        /// <returns>The resouce to retrieve</returns>
        public static object GetResource(ResourceDictionaryName resourceDictionaryName, object key)
        {
            var resourceDictionary = new ResourceDictionary
            {
                Source = new Uri(PathToResourceDictionaries + resourceDictionaryName.ToString() + ".xaml", UriKind.RelativeOrAbsolute)
            };

            return resourceDictionary[key];
        }
    }
}
