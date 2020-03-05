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
    public class RibbonControler : ExcelDna.Integration.CustomUI.ExcelRibbon
    {
        private string _ticker;
        private string _startDate;
        private string _endDate;

        public void OnDataImporterPressed(IRibbonControl control)
        {
            Excel.Application excel_application =
                (Excel.Application)ExcelDna.Integration.ExcelDnaUtil.Application;
           
            object selection = excel_application.Selection;
            if (selection is Excel.Range)
            {
                Excel.Range selected_range =
                    (Excel.Range)selection;
                int first_col = selected_range.Column;
                int first_row = selected_range.Row;

                selected_range.Worksheet.Cells[first_row, first_col].Value = _ticker;
                selected_range.Worksheet.Cells[first_row, first_col + 1].value = _startDate;
                selected_range.Worksheet.Cells[first_row, first_col + 2].value = _endDate;
            }
        }

        public void GetTickerValue(IRibbonControl control, string text=null)
        {
            if (!string.IsNullOrEmpty(text))
            {
                this._ticker = text;
            }
        }
        
        // TODO (FRZ): Add sanity check in input date variables
        public void GetStartDateValue(IRibbonControl control, string text=null)
        {
            if (!string.IsNullOrEmpty(text))
            {
                this._startDate = text;
            }
        }

        public void GetEndDateValue(IRibbonControl control, string text=null)
        {
            if (!string.IsNullOrEmpty(text))
            {
                this._endDate = text;
            }
        }
    }
}
