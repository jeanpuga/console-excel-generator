using System;
using OfficeOpenXml;
using Excel.Contracts;

namespace Excel.Cell
{
    public class CellHour : ICell
    {
        DateTime? Date;

        public CellHour(DateTime? date)
        {
            this.Date = date;
        }

        public void ApllyCell(ExcelRange cells)
        {
            cells.Style.Numberformat.Format = "HH:mm";
            cells.Value = this.Date;
        }
    }
}
