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
    /// <summary>
    /// EXCEL帮助类
    /// </summary>
    /// <typeparam name="T"></typeparam>
    public class ExcelHelper<T>
    {
        private int _startRow = 1;
        private int _startCol = 1;

        //sheet名
        private readonly string _sheetName;

        //数据
        private readonly List<T> _lines;

        //标题
        private List<List<string>> _titles;


        public ExcelHelper(string sheetName, List<T> lines)
        {
            _sheetName = sheetName;
            _lines = lines;
            _titles = new List<List<string>>(lines.Count);
        }

        /// <summary>
        /// 导出EXCEL流
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <returns></returns>
        public MemoryStream Export()
        {
            if (_sheetName.IsNullOrEmpty() || _lines?.Count == 0)
            {
                throw new ArgumentNullException($"{nameof(_sheetName)} or lines");
            }

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
        /// Description：标题行（内容来源于字段标题）
        /// Display：标题行(内容来源于字段值)
        /// </remarks>
        /// <param name="worksheet"></param>
        private void Aw_BuildHeaders(IXLWorksheet worksheet)
        {
            Aw_BuildTypeHeaders(typeof(T));
        }


        /// <summary>
        /// 构建类型表头信息
        /// </summary>
        /// <param name="type"></param>
        private List<string> Aw_BuildTypeHeaders(Type type)
        {
            var props = type.GetTypeProperties();

            var currentTitles = new List<string>();
            foreach (var prop in props)
            {
                var typeDesc = prop.PropertyType.ToString();
                if (prop.PropertyType.IsSimpleType())
                {
                    currentTitles.Add(Aw_BuildTitleHeader(prop));
                }
                else if(typeDesc.StartsWith("System.Collections.Generic.List"))
                {
                    Aw_BuildTypeHeaders(prop.PropertyType);
                }
                else
                {
                    Aw_BuildTypeHeaders(prop.PropertyType);
                }
            }

            return currentTitles;
        }


        /// <summary>
        /// 构建标题单元格内容
        /// </summary>
        /// <param name="prop"></param>
        /// <returns></returns>
        private string Aw_BuildTitleHeader(PropertyInfo prop)
        {
            var description = prop.GetCustomAttribute<DescriptionAttribute>()?.Description;
            if (description.IsNotNullOrEmpty())
            {
                return description;
            }

            //DisplayName
            var displayName = prop.GetCustomAttribute<DisplayNameAttribute>()?.DisplayName;
            if (displayName.IsNullOrEmpty())
            {
                return string.Empty;
            }

            if (prop.PropertyType.IsEnum)
            {
                return (prop.GetValue(_lines.FirstOrDefault()) as Enum).Desc();
            }

            return prop.GetValue(_lines.FirstOrDefault()).ToString();
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
