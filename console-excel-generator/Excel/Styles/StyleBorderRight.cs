
using Excel.Contracts;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace Excel.Styles
{
    public class StyleBorderRight : IStyle
    {
        ExcelBorderStyle BorderStyle; 

        public StyleBorderRight(ExcelBorderStyle borderStyle)
        {
            this.BorderStyle = borderStyle;
        }

        public void ApllyStyle(ExcelRange cells)
        {
            cells.Style.Border.Right.Style = this.BorderStyle;
        }
    }
}
