using OfficeOpenXml;
using Excel.Contracts;


namespace Excel.Cell
{
    public class CellPercentage : ICell
    {
        decimal? Value;
        decimal[] ValueArray;

        public CellPercentage(decimal? value)
        {
            this.Value = value;
        }

        public CellPercentage(decimal[] value)
        {
            this.ValueArray = value;
        }
        
        public void ApllyCell(ExcelRange cells)
        {
            cells.Style.Numberformat.Format = "0.0%";
            if (this.Value.HasValue)
            {
                cells.Value = this.Value;
            }
            else
            {
                cells.Value = this.ValueArray;
            }
        }

    }
}
