using OfficeOpenXml;

namespace Excel.Contracts
{
    public interface IStyle
    {
        void ApllyStyle(ExcelRange cells);
    }
}
