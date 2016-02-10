using System;
using System.Collections.Generic;
using System.ComponentModel.Composition;
using System.Linq;
using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.Common.Attribute
{
    [AttributeUsage(AttributeTargets.Class), MetadataAttribute]
    public class ExportImageRibbonIdAttribute : ExportAttribute, IRibbonIdMetadata
    {
        public ExportImageRibbonIdAttribute(params string[] ribbonIds)
            : base(typeof(ImageHandler))
        {
            RibbonIds = ribbonIds ?? Enumerable.Empty<string>();
        }

        public IEnumerable<string> RibbonIds { get; set; }
    }
}
