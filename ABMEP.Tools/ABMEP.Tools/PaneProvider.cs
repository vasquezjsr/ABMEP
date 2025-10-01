// ABMEP.Tools / PaneProvider.cs
using System.Windows;
using Autodesk.Revit.UI;

namespace ABMEP.Tools
{
    /// <summary>Thin provider that gives Revit our WPF element.</summary>
    public class PaneProvider : IDockablePaneProvider
    {
        private readonly FrameworkElement _element;

        public PaneProvider(FrameworkElement element) => _element = element;

        public void SetupDockablePane(DockablePaneProviderData data)
        {
            data.FrameworkElement = _element;
            // We set InitialState during registration; nothing else needed here.
        }
    }
}
