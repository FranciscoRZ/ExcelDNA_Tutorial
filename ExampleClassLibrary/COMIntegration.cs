using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using ExcelDna;
using ExcelDna.IntelliSense;
using System.Runtime.InteropServices;

namespace ExampleClassLibrary
{
    [ComVisible(false)]
    internal class ExcelAddin : ExcelDna.Integration.IExcelAddIn
    {
        public void AutoOpen()
        {
            ExcelDna.IntelliSense.IntelliSenseServer.Install();
        }
        public void AutoClose()
        {
            ExcelDna.IntelliSense.IntelliSenseServer.Uninstall();
        }
    }
}
