using GitHubActionsReportGenerator.SheetGenerators;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GitHubActionsReportGenerator.SheetUpdaters
{
    public class FlakyTestsSheetUpdater : BaseSheetGenerator
    {
        public FlakyTestsSheetUpdater()
        {
            SheetName = "Tests Likely To Be Flaky";
        }

        public override Task Update()
        {
            throw new NotImplementedException();
        }
    }
}
