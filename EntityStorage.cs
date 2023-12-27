
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;

namespace ExcelTask
{
    public class EntityStorage
    {
        public static readonly Dictionary<string, Type> StorageNames = new Dictionary<string, Type>()
        {
            ["Товары"] = typeof(Ware),
            ["Клиенты"] = typeof(Client),
            ["Заявки"] = typeof(Order)
        };

        private static EntityStorage instance;
        public static EntityStorage Instance => instance ?? (instance = new EntityStorage());
        private EntityStorage()
        {
            WareStorage = new List<Ware>();
            ClientStorage = new List<Client>();
            OrderStorage = new List<Order>();
        }

        public List<Ware> WareStorage { get; set; }
        public List<Client> ClientStorage { get; set; }
        public List<Order> OrderStorage { get; set; }

        public void SaveRowInStorage(Type type, Row row, WorkbookPart workbookPart)
        {
            var wrapper = new List<string>();
            FillWrapper(wrapper, row, workbookPart);

            if (type == typeof(Ware))
            {
                Ware storage = new Ware();
                storage.Id = int.TryParse(wrapper[0], out int result) ? result : default;
                storage.Name = wrapper[1];
                storage.Units = wrapper[2];
                storage.Price = decimal.TryParse(wrapper[3], out decimal result3) ? result3 : default;

                WareStorage.Add(storage);
            }

            if (type == typeof(Client))
            {
                Client storage = new Client();
                storage.Id = int.TryParse(wrapper[0], out int result) ? result : default;
                storage.Name = wrapper[1];
                storage.Address = wrapper[2];
                storage.Contact = wrapper[3];

                ClientStorage.Add(storage);
            }

            if (type == typeof(Order))
            {
                Order storage = new Order();
                storage.Id = int.TryParse(wrapper[0], out int result) ? result : default;
                storage.WareId = int.TryParse(wrapper[1], out int result1) ? result1 : default;
                storage.ClientId = int.TryParse(wrapper[2], out int result2) ? result2 : default;
                storage.Number = int.TryParse(wrapper[3], out int result3) ? result3 : default;
                storage.Quantity = int.TryParse(wrapper[4], out int result4) ? result4 : default;

                if (double.TryParse(wrapper[5], out double result5))
                    storage.Date = DateTime.FromOADate(result5);

                OrderStorage.Add(storage);
            }
        }

        private void FillWrapper(List<string> wrapper, Row row, WorkbookPart workbookPart)
        {
            foreach (Cell cell in row)
            {
                var value = cell.InnerText;
                if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
                {
                    var stringTable = workbookPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();
                    value = stringTable?.SharedStringTable?.ElementAt(int.Parse(value))?.InnerText ?? value;
                }
                wrapper.Add(value);
            }
        }

    }
}
