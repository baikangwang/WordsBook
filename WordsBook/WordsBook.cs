using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using WordsBook.Properties;
using Excel = Microsoft.Office.Interop.Excel;

namespace WordsBook
{
    public partial class WordsBook
    {
        private void WordsBook_Load(object sender, RibbonUIEventArgs e)
        {
            btnRecite.Checked = false;
            btnRecite.Image = Resources.p32;
        }

        private void btnRecite_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Window window = e.Control.Context;
            Excel.Worksheet activeWorksheet = ((Excel.Worksheet)window.Application.ActiveSheet);
            Excel.Range firstRow = activeWorksheet.Range["A1"];
            firstRow.EntireColumn.Hidden = btnRecite.Checked;
            btnRecite.Image = btnRecite.Checked ? Resources.ph32 : Resources.p32;
        }
    }
}
