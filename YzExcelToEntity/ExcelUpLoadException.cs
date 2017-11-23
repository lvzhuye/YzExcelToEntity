using System;

namespace YzExcelToEntity
{

    public class ExcelUpLoadException : Exception
    {
        public int ExcelRowNumber { get; set; }

        public int ExcelCellNumber { get; set; }

        public ExcelUpLoadException(int excelRowNumer, int excelCellNumber, string message)
            : base(message)
        {
            this.ExcelRowNumber = excelRowNumer;
            this.ExcelCellNumber = excelCellNumber;
        }
    }
}
