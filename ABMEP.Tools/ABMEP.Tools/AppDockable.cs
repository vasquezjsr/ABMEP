// File: ABMEP.Tools/AppDockable.cs
// Target: .NET Framework 4.8

using System;
using System.IO;
using System.Reflection;
using System.Windows.Controls;
using Autodesk.Revit.UI;

namespace ABMEP.Tools
{
    public class AppDockable : IExternalApplication
    {
        private static bool _registered = false;
        private static bool _contentLoaded = false;
        private static bool _shownOnce = false;

        private static readonly string HotloadDir = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            @"Autodesk\Revit\Addins\2024\ABMEP_Hotload");

        private const string WorkerFileName = "ABMEP.Work.dll";
        private const string SpoolPaneFullName = "ABMEP.Work.Views.SpoolPane";

        public Result OnStartup(UIControlledApplication app)
        {
            // 1) Register dockable pane once
            RegisterPane(app);

            // 2) Load UI once (with null UIApplication) and show the pane once to avoid the black surface
            SafeLoadPaneOnce(app, showNow: true);

            // No event subscriptions; we never force-show again.
            return Result.Succeeded;
        }

        public Result OnShutdown(UIControlledApplication app) => Result.Succeeded;

        // ---------------- helpers ----------------

        private static void RegisterPane(UIControlledApplication app)
        {
            if (_registered) return;

            var host = PaneHost.Instance;
            var provider = new PaneProvider(host);

            var state = new DockablePaneState
            {
                DockPosition = DockPosition.Tabbed,
                TabBehind = DockablePanes.BuiltInDockablePanes.ProjectBrowser
            };

            var data = new DockablePaneProviderData
            {
                FrameworkElement = host,
                InitialState = state,
                VisibleByDefault = true
            };

            try { app.RegisterDockablePane(PaneIds.AbmepPaneId, "ABMEP Assembly Manager", provider); }
            catch { /* already registered */ }

            _registered = true;
        }

        private static void SafeLoadPaneOnce(UIControlledApplication app, bool showNow)
        {
            try
            {
                var pane = app.GetDockablePane(PaneIds.AbmepPaneId);

                if (!_contentLoaded)
                {
                    // At app startup we don't have a UIApplication; construct with null.
                    var ui = TryCreateSpoolPaneUI(null);
                    if (ui != null)
                    {
                        PaneHost.Instance.SetContent(ui);
                        _contentLoaded = true;
                    }
                }

                if (showNow && !_shownOnce)
                {
                    try { pane.Show(); } catch { }
                    _shownOnce = true;
                }
            }
            catch
            {
                // If hotload files aren't there yet, nothing breaks. User can load via ribbon button later.
            }
        }

        /// <summary>
        /// Creates ABMEP.Work.Views.SpoolPane(UIApplication) from hotload DLL.
        /// Pass null when no UIApplication is available at startup.
        /// </summary>
        private static UserControl TryCreateSpoolPaneUI(UIApplication uiappOrNull)
        {
            try
            {
                string workerPath = Path.Combine(HotloadDir, WorkerFileName);
                if (!File.Exists(workerPath)) return null;

                byte[] dllBytes = File.ReadAllBytes(workerPath);
                string pdbPath = Path.ChangeExtension(workerPath, ".pdb");
                byte[] pdbBytes = File.Exists(pdbPath) ? File.ReadAllBytes(pdbPath) : null;

                Assembly asm = (pdbBytes != null)
                    ? Assembly.Load(dllBytes, pdbBytes)
                    : Assembly.Load(dllBytes);

                var type = asm.GetType(SpoolPaneFullName, throwOnError: false, ignoreCase: false);
                if (type == null) return null;

                // SpoolPane ctor: SpoolPane(UIApplication uiApp). We can pass null safely at startup.
                var ui = Activator.CreateInstance(type, uiappOrNull) as UserControl;
                return ui;
            }
            catch
            {
                return null;
            }
        }
    }
}
