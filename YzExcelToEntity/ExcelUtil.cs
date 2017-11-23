using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;

namespace YzExcelToEntity
{
    public static class ExcelUtil
    {
        /// <summary>
        /// 根据文件流，数据开始行，映射数据到数据
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="fs"></param>
        /// <param name="startDataRow">基于0开始</param>
        /// <returns></returns>
        public static List<T> ExcelFileStreamToEntities<T>(string filePath, int startDataRow) where T : ExcelMapBase, new()
        {
            FileStream fs = null;

            List<T> models = new List<T>();
            try
            {
                fs = new FileStream(filePath, FileMode.Open, FileAccess.Read);

                IWorkbook workbook = null;
                ISheet sheet = null;
                IRow cRow = null;

                if (filePath.Last() == 's')
                {

                    workbook = new HSSFWorkbook(fs);
                }
                else
                {
                    workbook = new XSSFWorkbook(fs);
                }

                sheet = (ISheet)workbook.GetSheetAt(0);
                int rowCount = sheet.LastRowNum;


                for (int j = startDataRow; j <= rowCount; j++)
                {
                    cRow = (IRow)sheet.GetRow(j);
                    T model = null;
                    try
                    {
                        model = getRowModel<T>(cRow);
                    }
                    catch (ExcelUpLoadException eule)
                    {
                        throw;
                    }
                    model.ExcelRowNum = j;
                    models.Add(model);
                }
            }
            catch (Exception ex)
            {
                throw;
            }
            finally
            {
                if (fs != null)
                    fs.Close();
            }

            return models;

        }

        //行转实体
        private static T getRowModel<T>(IRow row) where T : ExcelMapBase, new()
        {
            T model = new T();
            Type modelType = typeof(T);
            PropertyInfo[] propertyInfos = modelType.GetProperties();
            foreach (PropertyInfo propertyInfo in propertyInfos)
            {

                ExcelColumnAttribute excelColumnAtt = (ExcelColumnAttribute)propertyInfo.GetCustomAttributes(typeof(ExcelColumnAttribute), false).SingleOrDefault();
                if (excelColumnAtt != null)
                {
                    try
                    {
                        object cellValue = Convert.ChangeType(getExcelValue(row,excelColumnAtt.ColumnId), propertyInfo.PropertyType);
                        propertyInfo.SetValue(model, cellValue, null);
                    }
                    catch (Exception ex)
                    {
                        throw new ExcelUpLoadException(0, excelColumnAtt.ColumnId, "单元格格式异常");
                    }
                }

            }

            return model;
        }

        /// <summary>
        /// 根据单元格类型，转换为字符串
        /// </summary>
        /// <param name="row"></param>
        /// <param name="cellNum"></param>
        /// <returns></returns>
        public static string getExcelValue(IRow row,int cellNum)
        {
            if(row == null)
            {
                throw new ArgumentNullException("getExcelValue方法参数row为null");
            }
            var cell = row.GetCell(cellNum);
            if(cell == null)
            {
                throw new Exception("根据CellNum找不到单元格");
            }
            if (cell.CellType == CellType.String)
            {
                return cell.StringCellValue;
            }
            else if(cell.CellType == CellType.Numeric)
            {
                return cell.NumericCellValue.ToString();
            }
            else
            {
                throw new Exception("暂不支持除文本、数字以外单元格样式");
            }
        }
    }
}
