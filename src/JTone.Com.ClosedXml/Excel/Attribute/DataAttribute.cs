using System;


namespace JTone.Com.ClosedXml.Excel.Attribute
{
    /// <summary>
    /// 数据类型
    /// </summary>
    public class DataAttribute : System.Attribute
    {
        /// <summary>
        /// 数据类型
        /// </summary>
        public Type DataType { get;}

        public DataAttribute(Type dataType)
        {
            DataType = dataType;
        }
    }
}
