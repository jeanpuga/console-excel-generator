using OfficeOpenXml;
using Excel.Contracts;


namespace Excel.Cell
{
    public class CellMoney : ICell
    {
        decimal? Value;
        ICell[] ValueArray;

        public CellMoney(decimal? value)
        {
            this.Value = value;
        }

        public CellMoney(decimal[] value)
        {
            //ICell[] icell = new CellMoney[]();

            this.ValueArray = new CellMoney[value.Length];

            for (int i = 0; i < value.Length; i++)
            {
                this.ValueArray[i] = new CellMoney (value[i] );
            }
        }

        public void ApllyCell(ExcelRange cells)
        {
            cells.Style.Numberformat.Format = "#,##0.00";
            if (this.Value.HasValue) {
                cells.Value = this.Value;
            }
            else{
                //esta melhorando
                cells.Value = this.ValueArray;
            }
            
        }

        
    }
}
