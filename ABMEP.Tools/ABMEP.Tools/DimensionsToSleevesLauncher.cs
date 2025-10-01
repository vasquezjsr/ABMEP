// Target: .NET Framework 4.8
// Assembly: ABMEP.Tools.dll

using Autodesk.Revit.Attributes;

namespace ABMEP.Tools
{
    [Transaction(TransactionMode.Manual)]
    public class DimensionsToSleevesLauncher : HotloadCommandBase
    {
        // This must exactly match the worker in ABMEP.Work
        protected override string[] WorkerTypeNames => new[]
        {
            "ABMEP.Work.Commands.DimensionsToSleeves"
        };
    }
}
