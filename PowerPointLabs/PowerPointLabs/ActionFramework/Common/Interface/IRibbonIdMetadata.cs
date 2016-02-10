using System.Collections.Generic;

namespace PowerPointLabs.ActionFramework.Common.Interface
{
    public interface IRibbonIdMetadata
    {
        IEnumerable<string> RibbonIds { get; }
    }
}
