using System;
using System.Reflection;

namespace ExcelTask
{
    public static class ExcelHelper
    {
        public static void SetPropertyValue(IExcelEntity entity, PropertyInfo property, string newValue)
        {
            var type = property.PropertyType;

            if (type == typeof(int))
                property.SetValue(entity, int.TryParse(newValue, out int result) ? result : default);

            else if (type == typeof(decimal))
                property.SetValue(entity, decimal.TryParse(newValue, out decimal result3) ? result3 : default);

            else if (type == typeof(DateTime))
                property.SetValue(entity, double.TryParse(newValue, out double result5) ? DateTime.FromOADate(result5) : default);

            else property.SetValue(entity, newValue);
        }

        public static string GetExcelSheetNameByType(Type type)
        {
            var attr = type.GetCustomAttribute<ExcelSheetAttribute>();
            return attr != null ? attr.Caption : null;
        }
    }
}
