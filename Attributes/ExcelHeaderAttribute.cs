using System;

namespace ExcelTask
{
    [AttributeUsage(AttributeTargets.Property)]
    public class ExcelHeaderAttribute : Attribute
    {
        public string Caption { get; set; }

        public ExcelHeaderAttribute(string caption) { Caption = caption; }
    }
}
