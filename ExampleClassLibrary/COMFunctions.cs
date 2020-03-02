using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices;

namespace ExampleClassLibrary
{
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.AutoDual)]
    public class DataPair
    {
        public string Label { get; set; }
        public int LabelID { get; set; }
    }

    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.AutoDual)]
    public class ExampleClassLibraryCOMFunctions
    {
        public DataPair GetDataPair(string label, int labelid)
        {
            return new DataPair() { Label = label, LabelID = labelid };
        }

        public DateTime ThirtyDaysAgo()
        {
            return DateTime.Today - TimeSpan.FromDays(30);
        }
    }
}
