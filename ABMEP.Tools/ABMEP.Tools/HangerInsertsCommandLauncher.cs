﻿// Target: .NET Framework 4.8 | AnyCPU/x64
// Assembly: ABMEP.Tools.dll

using Autodesk.Revit.Attributes;

namespace ABMEP.Tools
{
    [Transaction(TransactionMode.Manual)]
    public class HangerInsertsCommandLauncher : HotloadCommandBase
    {
        // Matches the worker below (ABMEP.Work.HangerInsertsCommand)
        protected override string[] WorkerTypeNames => new[]
        {
            "ABMEP.Work.HangerInsertsCommand"
        };
    }
}
