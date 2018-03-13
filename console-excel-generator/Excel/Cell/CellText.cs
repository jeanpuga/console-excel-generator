using OfficeOpenXml;
using Excel.Contracts;


namespace Excel.Cell
{
    public class CellText : ICell
    {
        string Value;

        public CellText(string value)
        {
            this.Value = value;
        }

        public void ApllyCell(ExcelRange cells)
        {
            cells.Value = this.Value;
        }
    }
}
