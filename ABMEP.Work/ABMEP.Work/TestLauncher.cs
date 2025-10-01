// Target: .NET Framework 4.8 | Platform: x64 (or AnyCPU)
// Project: ABMEP.Tools  (build -> ABMEP.Tools.dll)
using System;
using System.IO;
using System.Linq;
using System.Reflection;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace ABMEP.Tools
{
    [Transaction(TransactionMode.Manual)]
    public class TestLauncher : IExternalCommand
    {
        // Same style as your working launchers
        private static readonly string HotloadDir =
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                         "Autodesk", "Revit", "Addins", "2024", "ABMEP_Hotload");

        private const string WorkerFileName = "ABMEP.Work.dll";
        private const string WorkerFullClassName = "ABMEP.Work.Test"; // implements IExternalCommand

        public Result Execute(ExternalCommandData c, ref string message, ElementSet elements)
        {
            try
            {
                Directory.CreateDirectory(HotloadDir);

                // Prefer newest timestamped worker if you’re dropping ABMEP.Work_*.dll via post-build
                string workerPath =
                    Directory.EnumerateFiles(HotloadDir, "ABMEP.Work_*.dll", SearchOption.TopDirectoryOnly)
                        .OrderByDescending(File.GetCreationTimeUtc)
                        .FirstOrDefault();

                if (workerPath == null)
                {
                    // Fallback to plain ABMEP.Work.dll
                    workerPath = Path.Combine(HotloadDir, WorkerFileName);
                    if (!File.Exists(workerPath))
                    {
                        TaskDialog.Show("Hotloader", $"No worker DLL found in:\n{HotloadDir}");
                        return Result.Cancelled;
                    }

                    // Copy to a unique temp file so the original never gets locked
                    string tempDir = Path.Combine(Path.GetTempPath(), "ABMEP_Hotload");
                    Directory.CreateDirectory(tempDir);
                    string tempDll = Path.Combine(
                        tempDir, $"{Path.GetFileNameWithoutExtension(workerPath)}_{DateTime.UtcNow:yyyyMMdd_HHmmssfff}.dll");
                    File.Copy(workerPath, tempDll, true);
                    workerPath = tempDll;
                }

                // Load the worker
                Assembly asm = Assembly.LoadFrom(workerPath);
                Type t = asm.GetType(WorkerFullClassName, throwOnError: false);
                if (t == null)
                {
                    TaskDialog.Show("Hotloader", $"Type not found:\n{WorkerFullClassName}");
                    return Result.Cancelled;
                }

                if (!(Activator.CreateInstance(t) is IExternalCommand cmd))
                {
                    TaskDialog.Show("Hotloader", $"Type does not implement IExternalCommand:\n{WorkerFullClassName}");
                    return Result.Cancelled;
                }

                // Optional: cleanup old timestamped versions (keep newest 10)
                try
                {
                    var old = Directory.EnumerateFiles(HotloadDir, "ABMEP.Work_*.dll")
                                       .OrderByDescending(File.GetCreationTimeUtc)
                                       .Skip(10)
                                       .ToList();
                    foreach (var f in old) { try { File.Delete(f); } catch { } }
                }
                catch { /* ignore */ }

                return cmd.Execute(c, ref message, elements);
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Hotloader", "Hotload failed:\n" + ex);
                return Result.Failed;
            }
        }
    }
}
