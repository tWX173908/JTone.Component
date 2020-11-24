using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Reflection;
using ClosedXML.Excel;
using JTone.Core;


namespace JTone.Com.ClosedXml.Excel
{
    public class ExcelHelper<T>
    {
        private int _startRow = 1;
        private int _startCol = 1;

        private readonly string _sheetName;
        private readonly List<T> _lines;

        public ExcelHelper(string sheetName, List<T> lines)
        {
            _sheetName = sheetName;
            _lines = lines;
        }

        /// <summary>
        /// 导出EXCEL流
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <returns></returns>
        public MemoryStream Export()
        {
            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add(_sheetName);

                Aw_BuildHeaders(worksheet);

                Aw_BuildLines(worksheet);

                var stream = new MemoryStream();
                workbook.SaveAs(stream);
                return stream;
            }
        }


        /// <summary>
        /// 构建标题栏
        /// </summary>
        /// <remarks>
        /// Display：标题行(内容来源于值)
        /// Description：标题行（内容来源于标题）
        /// </remarks>
        /// <param name="worksheet"></param>
        private void Aw_BuildHeaders(IXLWorksheet worksheet)
        {
            var data = _lines.FirstOrDefault();
            var props = typeof(T).GetTypeProperties();

            foreach (var prop in props)
            {
                Aw_BuildTypeHeaders(worksheet, prop, data);

                var description = prop.GetCustomAttribute<DescriptionAttribute>()?.Description;
                if (description.IsNotNullOrEmpty())
                {
                    worksheet.Cell(_startRow, _startCol++).Value = description;
                    continue;
                }

                //DisplayName
                var displayName = prop.GetCustomAttribute<DisplayNameAttribute>()?.DisplayName;


            }
        }


        /// <summary>
        /// 构建类型表头信息
        /// </summary>
        /// <param name="worksheet"></param>
        /// <param name="prop"></param>
        /// <param name="data"></param>
        private static void Aw_BuildTypeHeaders(IXLWorksheet worksheet, PropertyInfo prop, T data)
        {
            var type = prop.PropertyType.ToString();
            if (type.StartsWith("System.Collections.Generic.List")) 
            {

            }
            else
            {
                
            }
        }


        /// <summary>
        /// 构建数据行
        /// </summary>
        /// <typeparam name="TData"></typeparam>
        /// <param name="workbook"></param>
        private static void Aw_BuildLines(IXLWorksheet workbook)
        {

        }
    }
}
