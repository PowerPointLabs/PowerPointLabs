using System;
using System.Collections.Generic;
using System.ComponentModel.Composition;
using System.Linq;

using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.Common.Attribute
{
    /// <summary>
    /// Apply this attribute with Ribbon Id to CheckBoxActionHandler to register its factory
    /// </summary>
    [AttributeUsage(AttributeTargets.Class), MetadataAttribute]
    public class ExportCheckBoxActionRibbonIdAttribute : ExportAttribute, IRibbonIdMetadata
    {
        public ExportCheckBoxActionRibbonIdAttribute(params string[] ribbonIds)
            : base(typeof(CheckBoxActionHandler))
        {
            RibbonIds = ribbonIds ?? Enumerable.Empty<string>();
        }

        public IEnumerable<string> RibbonIds { get; set; }
    }
}
