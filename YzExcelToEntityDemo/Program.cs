using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using YzExcelToEntity;
using YzExcelToEntityDemo.Model;

namespace YzExcelToEntityDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            //文件路径地址
            string filePath = @"";
            List<TestModel> testModels = ExcelUtil.ExcelFileStreamToEntities<TestModel>(filePath, 1);
        }
    }
}
