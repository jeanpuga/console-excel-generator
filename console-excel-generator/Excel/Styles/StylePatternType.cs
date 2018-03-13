
using Excel.Contracts;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace Excel.Styles
{
    public class StylePatternType : IStyle
    {
        ExcelFillStyle FillStyle;

        public StylePatternType(ExcelFillStyle fillStyle)
        {
            this.FillStyle = fillStyle;
        }

        public void ApllyStyle(ExcelRange cells)
        {
            cells.Style.Fill.PatternType = this.FillStyle;
        }
    }
}
