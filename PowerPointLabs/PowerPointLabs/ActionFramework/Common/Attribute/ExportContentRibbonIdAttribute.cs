using System;
using System.Collections.Generic;
using System.ComponentModel.Composition;
using System.Linq;
using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.Common.Attribute
{
    /// <summary>
    /// Apply this attribute with Ribbon Id to ContentHandler to register its factory
    /// </summary>
    [AttributeUsage(AttributeTargets.Class), MetadataAttribute]
    public class ExportContentRibbonIdAttribute : ExportAttribute, IRibbonIdMetadata
    {
        public ExportContentRibbonIdAttribute(params string[] ribbonIds)
            : base(typeof(ContentHandler))
        {
            RibbonIds = ribbonIds ?? Enumerable.Empty<string>();
        }

        public IEnumerable<string> RibbonIds { get; set; }
    }
}
