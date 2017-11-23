using YzExcelToEntity;

namespace YzExcelToEntityDemo.Model
{
    public class TestModel:ExcelMapBase
    {
        [ExcelColumn(ColumnId =0)]
        public string Name { get; set; }

        [ExcelColumn(ColumnId =1)]
        public int Age { get; set; }
    }
}
