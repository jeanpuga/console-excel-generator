using Excel.Contracts;
using OfficeOpenXml;


namespace Excel.Styles
{
    public class StyleFontSize : IStyle
    {
        int Value;

        public StyleFontSize(int value)
        {
            this.Value = value;
        }

        public void ApllyStyle(ExcelRange cells)
        {
            cells.Style.Font.Size = this.Value;
        }
    }
}
