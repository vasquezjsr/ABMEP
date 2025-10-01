using System;
using System.IO;
using System.Reflection;
using System.Windows.Controls;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace ABMEP.Tools
{
    [Transaction(TransactionMode.Manual)]
    public class SpoolPaneLauncher : IExternalCommand
    {
        private static readonly string HotloadDir =
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                @"Autodesk\Revit\Addins\2024\ABMEP_Hotload");

        private const string WorkerFileName = "ABMEP.Work.dll";
        private const string SpoolPaneFullName = "ABMEP.Work.Views.SpoolPane";

        public Result Execute(ExternalCommandData c, ref string message, ElementSet elements)
        {
            try
            {
                var uiapp = c.Application;
                Directory.CreateDirectory(HotloadDir);
                var workerPath = Path.Combine(HotloadDir, WorkerFileName);
                if (!File.Exists(workerPath))
                {
                    TaskDialog.Show("ABMEP", $"Worker not found:\n{workerPath}");
                    return Result.Cancelled;
                }

                // Load ABMEP.Work from bytes (your existing hotload pattern)
                var bytes = File.ReadAllBytes(workerPath);
                var asm = Assembly.Load(bytes);

                // Create the SpoolPane(UserControl), pass UIApplication to ctor
                var paneType = asm.GetType(SpoolPaneFullName, throwOnError: true);
                var ui = (UserControl)Activator.CreateInstance(paneType, uiapp);

                // Inject into our dockable host, show pane
                var pane = uiapp.GetDockablePane(PaneIds.AbmepPaneId);
                ABMEP.Tools.PaneHost.Instance.SetContent(ui);
                pane.Show();

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                TaskDialog.Show("ABMEP", "Spool Pane load error:\r\n" + ex);
                return Result.Failed;
            }
        }
    }
}
