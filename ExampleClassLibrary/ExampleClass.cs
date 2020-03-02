using System;
using System.Windows;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices;

using ExcelDna;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Core;


namespace ExampleClassLibrary
{
    [ComVisible(true)]
    public class ExampleClass : ExcelDna.Integration.CustomUI.ExcelRibbon
    {
        private string _myValue;
        public void OnExample1ButtonPressed(IRibbonControl control)
        {
            MessageBoxResult mbr = MessageBox.Show("Enter Yes or No into selected Cell?",
                                                    "Choose",
                                                    MessageBoxButton.YesNo);
            Excel.Application excel_application =
                (Excel.Application)ExcelDna.Integration.ExcelDnaUtil.Application;

            string result = (mbr == MessageBoxResult.Yes) ? "Yes" : "No";
            
            object selection = excel_application.Selection;
            if (selection is Excel.Range)
            {
                Excel.Range selected_range =
                    (Excel.Range)selection;
                int first_col = selected_range.Column;
                int first_row = selected_range.Row;

                selected_range.Worksheet.Cells[first_row, first_col].Value = _myValue;
            }
        }

        public void GetEditBoxValue(IRibbonControl control, string text=null)
        {
            if (!string.IsNullOrEmpty(text))
            {
                _myValue = text;
            }
        }
    }
}
