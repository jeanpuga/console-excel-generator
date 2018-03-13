
using Excel.Contracts;
using OfficeOpenXml;
using OfficeOpenXml.Style;


namespace Excel.Styles
{
    public class StyleBorderLeft : IStyle
    {
        ExcelBorderStyle BorderStyle;

        public StyleBorderLeft(ExcelBorderStyle borderStyle)
        {
            this.BorderStyle = borderStyle;
        }

        public void ApllyStyle(ExcelRange cells)
        {
            cells.Style.Border.Left.Style = this.BorderStyle;
        }
    }
}
