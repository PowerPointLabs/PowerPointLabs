using System;
using System.Collections.Generic;
using System.ComponentModel.Composition;
using System.Linq;
using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.Common.Attribute
{
    [AttributeUsage(AttributeTargets.Class), MetadataAttribute]
    public class ExportSupertipRibbonIdAttribute : ExportAttribute, IRibbonIdMetadata
    {
        public ExportSupertipRibbonIdAttribute(params string[] ribbonIds)
            : base(typeof(SupertipHandler))
        {
            RibbonIds = ribbonIds ?? Enumerable.Empty<string>();
        }

        public IEnumerable<string> RibbonIds { get; set; }
    }
}
