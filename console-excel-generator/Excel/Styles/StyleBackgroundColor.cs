using Excel.Contracts;
using OfficeOpenXml;
using System.Drawing;

namespace Excel.Styles
{
    public class StyleBackgroundColor : IStyle
    {
        string HexColor;

        public StyleBackgroundColor(string hexColor)
        {
            this.HexColor = hexColor;
        }

        public void ApllyStyle(ExcelRange cells)
        {
            cells.Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml(this.HexColor));
        }
    }
}
