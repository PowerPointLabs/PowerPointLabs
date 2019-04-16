using System;
using System.Collections.Generic;
using System.ComponentModel.Composition;
using System.Linq;

using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.Common.Attribute
{
    /// <summary>
    /// Apply this attribute with Ribbon Id to LabelHandler to register its factory
    /// </summary>
    [AttributeUsage(AttributeTargets.Class), MetadataAttribute]
    public class ExportLabelRibbonIdAttribute : ExportAttribute, IRibbonIdMetadata
    {
        public ExportLabelRibbonIdAttribute(params string[] ribbonIds)
            : base(typeof(LabelHandler))
        {
            RibbonIds = ribbonIds ?? Enumerable.Empty<string>();
        }

        public IEnumerable<string> RibbonIds { get; set; }
    }
}
