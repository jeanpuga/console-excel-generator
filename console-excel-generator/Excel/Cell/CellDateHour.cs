using Excel.Contracts;
using OfficeOpenXml;
using System;

namespace Excel.Cell
{
    public class CellDateHour : ICell
    {
        DateTime? Date;

        public CellDateHour(DateTime? date)
        {
            this.Date = date;
        }

        public void ApllyCell(ExcelRange cells)
        {
            cells.Style.Numberformat.Format = "dd/MM/yyyy HH:mm";
            cells.Value = this.Date;
        }
    }
}
