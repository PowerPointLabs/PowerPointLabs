using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.AgendaLab
{
    [ExportLabelRibbonId(TextCollection.BeamAgendaTag)]
    class BeamAgendaLabelHandler : LabelHandler
    {
        protected override string GetLabel(string ribbonId)
        {
            return TextCollection.AgendaLabBeamAgendaButtonLabel;
        }
    }
}
