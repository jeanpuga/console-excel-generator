using OfficeOpenXml.Style;



namespace Excel.Contracts
{
    public interface IStyleFactory
    {
        IStyle CreateStyleFontSize(int value);
        IStyle CreateStyleFontColor(string hexColor);
        IStyle CreateStyleFontBold();
        IStyle CreateStylePatternType(ExcelFillStyle fillStyle);
        IStyle CreateStyleBackgroundColor(string hexColor);
        IStyle CreateStyleIndent(int value);
        IStyle CreateStyleHorizontalAlignment(ExcelHorizontalAlignment horizontalAlignment);
        IStyle CreateStyleBorderLeft(ExcelBorderStyle borderStyle);
        IStyle CreateStyleBorderRight(ExcelBorderStyle borderStyle);
        IStyle CreateStyleBorderTop(ExcelBorderStyle borderStyle);
        IStyle CreateStyleBorderBottom(ExcelBorderStyle borderStyle);
        IStyle CreateStyleMerge(bool isMerge);
    }
}