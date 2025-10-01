// Target: .NET Framework 4.8
using System;
using System.IO;
using System.Linq;
using System.Reflection;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace ABMEP.Tools
{
    /// <summary>Base for all launchers that hotload ABMEP.Work.* commands.</summary>
    public abstract class HotloadCommandBase : IExternalCommand
    {
        /// <summary>Return 1..N candidate fully-qualified worker types to try, e.g. "ABMEP.Work.PipeTrimbleCommand".</summary>
        protected abstract string[] WorkerTypeNames { get; }

        public Result Execute(ExternalCommandData c, ref string message, ElementSet elements)
        {
            try
            {
                string appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
                string hotloadDir = Path.Combine(appData, "Autodesk", "Revit", "Addins", "2024", "ABMEP_Hotload");
                string workerPath = Path.Combine(hotloadDir, "ABMEP.Work.dll");

                if (!File.Exists(workerPath))
                {
                    TaskDialog.Show("ABMEP Hotload", $"Worker not found:\n{workerPath}");
                    return Result.Cancelled;
                }

                using (new HotloadResolver(hotloadDir))
                {
                    byte[] dllBytes = File.ReadAllBytes(workerPath);
                    string pdbPath = Path.ChangeExtension(workerPath, ".pdb");
                    byte[] pdbBytes = File.Exists(pdbPath) ? File.ReadAllBytes(pdbPath) : null;

                    Assembly workAsm = (pdbBytes != null)
                        ? Assembly.Load(dllBytes, pdbBytes)
                        : Assembly.Load(dllBytes);

                    // Try all provided names in order; first one that exists wins.
                    Type t = null;
                    foreach (var name in WorkerTypeNames ?? Array.Empty<string>())
                    {
                        t = workAsm.GetType(name, throwOnError: false, ignoreCase: false);
                        if (t != null) break;
                    }
                    if (t == null)
                        throw new TypeLoadException("None of the worker types were found in ABMEP.Work.dll:\n" +
                            string.Join("\n", WorkerTypeNames ?? Array.Empty<string>()));

                    var cmd = (IExternalCommand)Activator.CreateInstance(t);
                    return cmd.Execute(c, ref message, elements);
                }
            }
            catch (Exception ex)
            {
                TaskDialog.Show("ABMEP Hotload Error", ex.ToString());
                return Result.Failed;
            }
        }

        /// <summary>Resolve any dependent DLLs from the hotload folder, from bytes (avoids file locks).</summary>
        private sealed class HotloadResolver : IDisposable
        {
            private readonly string _dir;
            public HotloadResolver(string dir)
            {
                _dir = dir;
                AppDomain.CurrentDomain.AssemblyResolve += OnResolve;
            }
            private Assembly OnResolve(object sender, ResolveEventArgs args)
            {
                var name = new AssemblyName(args.Name).Name + ".dll";
                var path = Path.Combine(_dir, name);
                if (!File.Exists(path)) return null;

                byte[] dll = File.ReadAllBytes(path);
                string pdbPath = Path.ChangeExtension(path, ".pdb");
                byte[] pdb = File.Exists(pdbPath) ? File.ReadAllBytes(pdbPath) : null;

                return (pdb != null) ? Assembly.Load(dll, pdb) : Assembly.Load(dll);
            }
            public void Dispose()
            {
                AppDomain.CurrentDomain.AssemblyResolve -= OnResolve;
            }
        }
    }
}
