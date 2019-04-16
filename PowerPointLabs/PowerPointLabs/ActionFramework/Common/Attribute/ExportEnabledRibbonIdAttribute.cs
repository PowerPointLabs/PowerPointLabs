using System;
using System.Collections.Generic;
using System.ComponentModel.Composition;
using System.Linq;

using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.Common.Attribute
{
    /// <summary>
    /// Apply this attribute with Ribbon Id to EnabledHandler to register its factory
    /// </summary>
    [AttributeUsage(AttributeTargets.Class), MetadataAttribute]
    public class ExportEnabledRibbonIdAttribute : ExportAttribute, IRibbonIdMetadata
    {
        public ExportEnabledRibbonIdAttribute(params string[] ribbonIds)
            : base(typeof(EnabledHandler))
        {
            RibbonIds = ribbonIds ?? Enumerable.Empty<string>();
        }

        public IEnumerable<string> RibbonIds { get; set; }
    }
}