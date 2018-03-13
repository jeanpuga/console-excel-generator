using Excel.Contracts;
using OfficeOpenXml;


namespace Excel.Styles
{
    public class StyleFontBold : IStyle
    {
        public void ApllyStyle(ExcelRange cells)
        {
            cells.Style.Font.Bold = true;
        }
    }
}
