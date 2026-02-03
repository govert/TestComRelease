using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;
using ExcelDna.Integration.Extensibility;
using ExcelDna.IntelliSense;
using System.Runtime.InteropServices;

namespace Test180
{
    internal class MyAddIn : ExcelRibbon, IExcelAddIn
    {
        #region IExcelAddin

        public void AutoOpen()
        {
            IntelliSenseServer.Install();
        }

        public void AutoClose()
        {
            IntelliSenseServer.Uninstall();
        }

        #endregion

        #region ExcelDna.Integration.CustomUI.ExcelRibbon

        //public override void OnAddInsUpdate(ref Array custom) { }

        //public override void OnBeginShutdown(ref Array custom) { }

        //public override void OnConnection(object Application, ext_ConnectMode ConnectMode, object AddInInst, ref Array custom) { }

        public override void OnDisconnection(ext_DisconnectMode RemoveMode, ref Array custom)
        {
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }

        //public override void OnStartupComplete(ref Array custom) { }

        #endregion
    }
}
