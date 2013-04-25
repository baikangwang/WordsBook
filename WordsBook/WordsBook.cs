using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using Excel = Microsoft.Office.Interop.Excel;

namespace WordsBook
{
    public partial class WordsBook
    {
        private void WordsBook_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void btnRecite_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Window window = e.Control.Context;
            Excel.Worksheet activeWorksheet = ((Excel.Worksheet)window.Application.ActiveSheet);
            Excel.Range firstRow = activeWorksheet.get_Range("A1");
            firstRow.EntireColumn.Hidden = btnRecite.Checked;
        }
    }
}
