using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace EliminateZeros1
{
    public partial class Ribbon
    {
        private void Ribbon_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Range Range = Globals.ThisAddIn.Application.ActiveSheet.UsedRange;

            Excel.Range usedRange = Globals.ThisAddIn.Application.Selection;

            if (usedRange != null)
            {
                var nRows = Range.Rows.Count;
                int nCols = usedRange.Columns.Count;
                for (int iRow = 1; iRow <= nRows; iRow++)
                {
                    for (int iCount = 1; iCount <= nCols; iCount++)
                    {
                        var cell = usedRange.Cells[iRow, iCount];
                        try
                        {
                            if (cell.Value == "" || cell.Value == null) continue;
                            cell.NumberFormat = "0";
                            cell.Value = (int)(Convert.ToDecimal(cell.Value) * 1);
                            //cell.Value = Int32.Parse(cell.Value) * 1;
                        }
                        catch
                        {
                            // MessageBox.Show("Error");
                        }
                    }
                }
            }
        }
    }
}
