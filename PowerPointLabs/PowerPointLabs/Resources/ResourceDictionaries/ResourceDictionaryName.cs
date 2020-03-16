using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerPointLabs.Resources.ResourceDictionaries
{
    /// <summary>
    /// The ResourceDictionaryName enum is meant to provide a safe way to refer to the Resource
    /// Dictionaries found in the ResourceDictionaries folder, as opposed to only using strings.
    /// </summary>
    public enum ResourceDictionaryName
    {
        // Notice to developers: Every value in this Enum must have its name match with exactly
        // one ResourceDictionary XAML file in the ResourceDictionaries folder (as given in
        // ResourceDictionaryUtil.PathToResourceDictionaries).
        GeneralResourceDictionary
    }
}
