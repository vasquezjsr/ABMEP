// Target: .NET Framework 4.8
// Assembly: ABMEP.Tools.dll

using Autodesk.Revit.Attributes;

namespace ABMEP.Tools
{
    [Transaction(TransactionMode.Manual)]
    public class StrutLengthFromWidthLauncher : HotloadCommandBase
    {
        protected override string[] WorkerTypeNames => new[]
        {
            "ABMEP.Work.StrutLengthFromWidthCommand"
        };
    }
}
