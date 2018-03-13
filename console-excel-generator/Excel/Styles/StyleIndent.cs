using Excel.Contracts;
using OfficeOpenXml;


namespace Excel.Styles
{
    public class StyleIndent : IStyle
    {
        int Value;

        public StyleIndent(int value)
        {
            this.Value = value;
        }

        public void ApllyStyle(ExcelRange cells)
        {
            cells.Style.Indent = this.Value;
        }
    }
}
