using Excel.Contracts;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace Excel.Styles
{
    public class StyleBorderBottom : IStyle
    {
        ExcelBorderStyle BorderStyle;

        public StyleBorderBottom(ExcelBorderStyle borderStyle)
        {
            this.BorderStyle = borderStyle;
        }

        public void ApllyStyle(ExcelRange cells)
        {
            cells.Style.Border.Bottom.Style = this.BorderStyle;
        }
    }
}
