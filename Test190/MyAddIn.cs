using ExcelDna.Integration;
using ExcelDna.IntelliSense;

namespace Test190
{
    internal class MyAddIn : IExcelAddIn
    {
        public void AutoOpen()
        {
            IntelliSenseServer.Install();
        }

        public void AutoClose()
        {
            IntelliSenseServer.Uninstall();
        }
    }
}
