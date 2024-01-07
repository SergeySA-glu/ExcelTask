using System;

namespace ExcelTask
{
    [AttributeUsage(AttributeTargets.Class)]
    public class ExcelSheetAttribute : Attribute
    {
        public string Caption { get; set; }

        public ExcelSheetAttribute(string caption) { Caption = caption; }
    }
}
