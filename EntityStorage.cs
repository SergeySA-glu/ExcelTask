
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text.RegularExpressions;

namespace ExcelTask
{
    public class EntityStorage
    {
        private static EntityStorage instance;
        public static EntityStorage Instance => instance ?? (instance = new EntityStorage());

        private readonly Dictionary<string, string> _headers = new Dictionary<string, string>();

        public Dictionary<Type, Dictionary<PropertyInfo, string>> ColumnCharTree { get; set; }

        public Dictionary<Type, List<object>> TypeStorages { get; set; }

        public string PathFile { get; set; }

        private EntityStorage()
        {
            ColumnCharTree = new Dictionary<Type, Dictionary<PropertyInfo, string>>();
            TypeStorages = new Dictionary<Type, List<object>>();
            foreach (var type in Assembly.GetExecutingAssembly().GetTypes().Where(t => t.GetInterfaces().Contains(typeof(IExcelEntity))))
                TypeStorages[type] = new List<object>();
        }

        public void SaveRowInStorage(IExcelEntity entity, Type type, Row row, WorkbookPart workbookPart)
        {
            var wrapper = new Dictionary<string, string>();
            FillWrapper(wrapper, row, workbookPart);
            entity.RowIndex = row.RowIndex;
            FillEntity(entity, wrapper);
            TypeStorages[type].Add(entity);
        }

        public void SetSheetHeaders(Row row, WorkbookPart workbookPart)
        {
            _headers.Clear();
            FillWrapper(_headers, row, workbookPart);
        }

        public void ClearStorage()
        {
            foreach (var type in TypeStorages.Keys)
                TypeStorages[type].Clear();

            ColumnCharTree.Clear();
        }

        private void FillWrapper(Dictionary<string, string> wrapper, Row row, WorkbookPart workbookPart)
        {
            foreach (Cell cell in row)
            {
                var value = cell.InnerText;
                if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
                {
                    var stringTable = workbookPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();
                    value = stringTable?.SharedStringTable?.ElementAt(int.Parse(value))?.InnerText ?? value;
                }
                wrapper.Add(Regex.Replace(cell.CellReference, @"[^A-Z]+", string.Empty), value);
            }
        }

        private void FillEntity(IExcelEntity entity, Dictionary<string, string> wrapper)
        {
            var type = entity.GetType();
            var properties = type.GetProperties();
            var propDict = new Dictionary<PropertyInfo, string>();
            foreach (var wrapValue in wrapper)
            {
                foreach (var property in properties)
                {
                    var headerAttr = property.GetCustomAttribute<ExcelHeaderAttribute>();
                    if (headerAttr != null)
                    {
                        if (_headers.ContainsKey(wrapValue.Key) && _headers[wrapValue.Key] == headerAttr.Caption)
                        {
                            if (!ColumnCharTree.ContainsKey(type))
                                propDict[property] = wrapValue.Key;
                            ExcelHelper.SetPropertyValue(entity, property, wrapValue.Value);
                        }
                    }
                }
            }

            if (!ColumnCharTree.ContainsKey(type))
                ColumnCharTree.Add(type, propDict);
        }
    }
}
