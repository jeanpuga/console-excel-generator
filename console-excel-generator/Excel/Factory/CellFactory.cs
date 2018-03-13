using System;
using Excel.Contracts;
using Excel.Cell;

namespace Excel.Factory
{
    public class CellFactory : ICellFactory
    {
        public ICell CreateCellDate(DateTime? value)
        {
            return new CellDate(value);
        }
        public ICell CreateCellDateHour(DateTime? value)
        {
            return new CellDateHour(value);
        }
        public ICell CreateCellHour(DateTime? value)
        {
            return new CellHour(value);
        }
        public ICell CreateCellMoney(decimal? value)
        {
            return new CellMoney(value);
        }
        public ICell CreateCellMoney(decimal[] value)
        {
            return new CellMoney(value);
        }
        public ICell CreateCellNumber(int? value)
        {
            return new CellNumber(value);
        }
        public ICell CreateCellPercentage(decimal? value)
        {
            return new CellPercentage(value);
        }
        public ICell CreateCellPercentage(decimal[] value)
        {
            return new CellPercentage(value);
        }
        public ICell CreateCellText(string value)
        {
            return new CellText(value);
        }
    }
}
