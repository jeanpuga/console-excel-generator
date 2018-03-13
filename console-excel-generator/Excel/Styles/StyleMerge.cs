using Excel.Contracts;
using OfficeOpenXml;

namespace Excel.Styles
{
    public class StyleMerge : IStyle
    {
        bool IsMerge;

        public StyleMerge(bool isMerge)
        {
            this.IsMerge = isMerge;
        }

        public void ApllyStyle(ExcelRange cells)
        {
            cells.Merge = this.IsMerge;
        }
    }
}
