using System;
using OfficeOpenXml;
using Excel.Contracts;

namespace Excel.Cell
{
    public class CellDate : ICell
    {
        DateTime? Date;

        public CellDate(DateTime? date)
        {
            this.Date = date;
        }

        public void ApllyCell(ExcelRange cells)
        {
            cells.Style.Numberformat.Format = "dd/MM/yyyy";
            cells.Value = this.Date;
        }
    }
}
