using System;

namespace Excel.Contracts
{
    public interface ICellFactory
    {
        ICell CreateCellText(string value);
        ICell CreateCellDate(DateTime? value);
        ICell CreateCellMoney(decimal? value);
        ICell CreateCellMoney(decimal[] value);

        ICell CreateCellNumber(int? value);
        ICell CreateCellPercentage(decimal? value);
        ICell CreateCellPercentage(decimal[] value);

        ICell CreateCellHour(DateTime? value);
        ICell CreateCellDateHour(DateTime? value);
    }
}
