using System;

namespace YzExcelToEntity
{
    [AttributeUsage(AttributeTargets.Property)]
    public class ExcelColumnAttribute : Attribute
    {
        private int _columnId;

        public int ColumnId
        {
            get
            {
                return _columnId;
            }
            set
            {
                _columnId = value;
            }
        }
    }
}
