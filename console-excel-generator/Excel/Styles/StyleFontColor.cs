
using Excel.Contracts;
using OfficeOpenXml;
using System.Drawing;

namespace Excel.Styles
{
    public class StyleFontColor : IStyle
    {
        string HexColor;

        public StyleFontColor(string hexColor)
        {
            this.HexColor = hexColor;
        }

        public void ApllyStyle(ExcelRange cells)
        {
            cells.Style.Font.Color.SetColor(ColorTranslator.FromHtml(this.HexColor));
        }
    }
}
