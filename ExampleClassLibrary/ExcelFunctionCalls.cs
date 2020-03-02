using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using ExcelDna;

namespace ExampleClassLibrary
{
    public static class ExcelFunctionCalls
    {
        [ExcelDna.Integration.ExcelFunction(Description = "Last week example UDF")]
        public static DateTime GetLastWeek()
        {
            return DateTime.Today;
        }

        [ExcelDna.Integration.ExcelFunction(Description = "Reverse string example UDF")]
        public static string ReverseString([ExcelDna.Integration.ExcelArgument(Name = "input_string",
                                                                               Description = "String to reverse")] 
                                            string input_string)
        {
            string reverse_string = "";
            int ilength = input_string.Length - 1;
            while (ilength > 0)
            {
                reverse_string = reverse_string + input_string[ilength];
                ilength--;
            }
            return reverse_string;
        }
    }
}
