// Target: .NET Framework 4.8 | AnyCPU/x64
// Assembly: ABMEP.Tools.dll
// Purpose: Launcher that hotloads ABMEP.Work.PipeTrimble via HotloadCommandBase.

using Autodesk.Revit.Attributes;

namespace ABMEP.Tools
{
    [Transaction(TransactionMode.Manual)]
    public class PipeTrimbleLauncher : HotloadCommandBase
    {
        protected override string[] WorkerTypeNames => new[]
        {
            "ABMEP.Work.PipeTrimble"
        };
    }
}
