// ABMEP.Tools / PaneIds.cs
using Autodesk.Revit.UI;
using System;

namespace ABMEP.Tools
{
    public static class PaneIds
    {
        // Use a stable GUID and never change it after first release.
        public static readonly DockablePaneId AbmepPaneId =
            new DockablePaneId(new Guid("9E7B0D86-4C1E-4F34-9F5B-9C0C0A2D5E11"));
    }
}
