
using Excel.Contracts;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace Excel.Styles
{
    public class StyleHorizontalAlignment : IStyle
    {
        ExcelHorizontalAlignment HorizontalAlignment;

        public StyleHorizontalAlignment(ExcelHorizontalAlignment horizontalAlignment)
        {
            this.HorizontalAlignment = horizontalAlignment;
        }

        public void ApllyStyle(ExcelRange cells)
        {
            cells.Style.HorizontalAlignment = this.HorizontalAlignment;
        }
    }
}
