using OfficeOpenXml;
using Excel.Contracts;


namespace Excel.Cell
{
    public class CellNumber : ICell
    {
        int? Value;

        public CellNumber(int? value)
        {
            this.Value = value;
        }

        public void ApllyCell(ExcelRange cells)
        {
            cells.Style.Numberformat.Format = "##0";
            cells.Value = this.Value;
        }
    }
}
