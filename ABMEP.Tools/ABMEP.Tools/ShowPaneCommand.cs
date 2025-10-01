using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace ABMEP.Tools
{
    [Transaction(TransactionMode.Manual)]
    public class ShowPaneCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData c, ref string message, ElementSet elements)
        {
            var pane = c.Application.GetDockablePane(PaneIds.AbmepPaneId);
            if (pane != null)
            {
                // Brings the pane up (if hidden). Revit 2024 has no Activate().
                pane.Show();
            }
            return Result.Succeeded;
        }
    }
}
