using System;

namespace OnReport
{
    /// <summary>
    /// Report DataSource definition
    /// </summary>
    public class OnReportDataSource
    {
        /// <summary>
        /// DataSet name
        /// </summary>
        public string Name { get; private set; }

        /// <summary>
        /// Report's data
        /// </summary>
        public object Value { get; private set; }

        public string ParameterName { get; internal set; }

        public OnReportDataSource(string name, object value)
        {
            Name = name;
            Value = value;
        }
    }
}
