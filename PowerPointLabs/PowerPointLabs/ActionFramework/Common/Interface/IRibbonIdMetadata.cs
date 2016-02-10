using System.Collections.Generic;

namespace PowerPointLabs.ActionFramework.Common.Interface
{
    /// <summary>
    /// Metadata used to describe which Ribbon Id to register
    /// </summary>
    public interface IRibbonIdMetadata
    {
        IEnumerable<string> RibbonIds { get; }
    }
}
