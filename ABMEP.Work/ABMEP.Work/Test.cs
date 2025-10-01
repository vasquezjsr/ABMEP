// Target: .NET Framework 4.8 | x64
// Assembly: ABMEP.Work.dll
using System;
using System.IO;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace ABMEP.Work
{
    [Transaction(TransactionMode.Manual)]
    public class Test : IExternalCommand
    {
        public Result Execute(ExternalCommandData c, ref string message, ElementSet elements)
        {
            // Put whatever you’re testing here. Change this text, rebuild, click button again.
            string appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            string hotloadDir = Path.Combine(appData, "Autodesk", "Revit", "Addins", "2024", "ABMEP_Hotload");
            string workerDll = Path.Combine(hotloadDir, "ABMEP.Work.dll");
            string ts = File.Exists(workerDll) ? File.GetLastWriteTime(workerDll).ToString("g") : "n/a";

            TaskDialog.Show("ABMEP Test Worker",
                "Hello from ABMEP.Work.Test\n\n" +
                $"Hotload DLL last write: {ts}\n" +
                "Hello Joe.");

            return Result.Succeeded;
        }
    }
}
