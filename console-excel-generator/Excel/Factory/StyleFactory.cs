
using OfficeOpenXml.Style;
using Excel.Styles;
using Excel.Contracts;

namespace Excel.Factory
{
    public class StyleFactory : IStyleFactory
    {
        public IStyle CreateStyleBackgroundColor(string hexColor)
        {
            return new StyleBackgroundColor(hexColor);
        }
        public IStyle CreateStyleFontBold()
        {
            return new StyleFontBold();
        }
        public IStyle CreateStyleFontColor(string hexColor)
        {
            return new StyleFontColor(hexColor);
        }
        public IStyle CreateStyleFontSize(int value)
        {
            return new StyleFontSize(value);
        }
        public IStyle CreateStyleIndent(int value)
        {
            return new StyleIndent(value);
        }
        public IStyle CreateStylePatternType(ExcelFillStyle fillStyle)
        {
            return new StylePatternType(fillStyle);
        }
        public IStyle CreateStyleHorizontalAlignment(ExcelHorizontalAlignment horizontalAlignment)
        {
            return new StyleHorizontalAlignment(horizontalAlignment);
        }
        public IStyle CreateStyleBorderLeft(ExcelBorderStyle borderStyle)
        {
            return new StyleBorderLeft(borderStyle);
        }
        public IStyle CreateStyleBorderRight(ExcelBorderStyle borderStyle)
        {
            return new StyleBorderRight(borderStyle);
        }
        public IStyle CreateStyleBorderTop(ExcelBorderStyle borderStyle)
        {
            return new StyleBorderTop(borderStyle);
        }
        public IStyle CreateStyleBorderBottom(ExcelBorderStyle borderStyle)
        {
            return new StyleBorderBottom(borderStyle);
        }
        public IStyle CreateStyleMerge(bool isMerge)
        {
            return new StyleMerge(isMerge);
        }
    }
}
