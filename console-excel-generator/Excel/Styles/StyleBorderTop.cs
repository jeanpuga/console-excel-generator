
using Excel.Contracts;
using OfficeOpenXml;
using OfficeOpenXml.Style;


namespace Excel.Styles
{
    public class StyleBorderTop : IStyle
    {
        ExcelBorderStyle BorderStyle;

        public StyleBorderTop(ExcelBorderStyle borderStyle)
        {
            this.BorderStyle = borderStyle;
        }

        public void ApllyStyle(ExcelRange cells)
        {
            cells.Style.Border.Top.Style = this.BorderStyle;
        }
    }
}
