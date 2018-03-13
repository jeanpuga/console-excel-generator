using OfficeOpenXml;


namespace Excel.Contracts
{
    public interface ICell
    {
        void ApllyCell(ExcelRange cells);
    }
}
