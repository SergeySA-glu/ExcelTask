using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;

namespace ExcelTask
{
    public class Executor
    {
        public Exception InnerException { get; set; }

        private bool _dataLoaded;


        private static Executor _instance;
        public static Executor Instance => _instance ?? (_instance = new Executor());

        public bool LoadEntities(string pathFile)
        {
            try
            {
                if (_dataLoaded)
                    throw new Exception("Данные уже были загружены");

                using (var document = SpreadsheetDocument.Open(pathFile, false))
                {
                    WorkbookPart workbookPart = document.WorkbookPart;
                    Workbook workbook = workbookPart.Workbook;
                    Sheets sheets = workbook.Elements<Sheets>().First();
                    foreach (var dict in EntityStorage.StorageNames)
                    {
                        var sheet = sheets.Elements<Sheet>().SingleOrDefault(s => s.Name == dict.Key);
                        var workSheetPart = (WorksheetPart)document.WorkbookPart.GetPartById(sheet.Id.Value);
                        var sheetData = workSheetPart.Worksheet.Elements<SheetData>().First();

                        foreach (var row in sheetData.Elements<Row>())
                        {
                            if (row.Equals(sheetData.Elements<Row>().First()))
                                continue;
                            EntityStorage.Instance.SaveRowInStorage(dict.Value, row, workbookPart);
                        }
                    }
                }
                _dataLoaded = true;
            }
            catch (Exception ex)
            {
                InnerException = ex;
                _dataLoaded = false;
            }
            return _dataLoaded;
        }

        public List<(DateTime, Client, int, decimal)> GetClientsByWare(string wareName)
        {
            var ordersInfo = new List<(DateTime, Client, int, decimal)>();
            try
            {
                var ware = EntityStorage.Instance.WareStorage.Where(w => w.Name == wareName).FirstOrDefault();
                if (ware == null)
                    throw new Exception("Товар не найден");

                var orders = EntityStorage.Instance.OrderStorage.Where(order => order.WareId == ware.Id);
                if (!orders.Any())
                    throw new Exception("Заказов по указанному товару не найдено");

                foreach (var order in orders)
                    ordersInfo.Add((order.Date, EntityStorage.Instance.ClientStorage.FirstOrDefault(cl => cl.Id == order.ClientId), order.Quantity, ware.Price));
            }
            catch (Exception ex)
            {
                InnerException = ex;
            }
            return ordersInfo;
        }

        public void SetNewContact(string organization, string newContact)
        {
            try
            {
                throw new NotImplementedException();
            }
            catch (Exception ex)
            {
                InnerException = ex;
            }
        }

        public Client GetGoldClient(bool year, int dateValue)
        {
            Client goldClient = null;
            try
            {
                var ordersByInterval = EntityStorage.Instance.OrderStorage.Where(order => year ? order.Date.Year == dateValue : order.Date.Month == dateValue && order.Date != default);
                var clientIds = ordersByInterval.Select(order => order.ClientId).Distinct();

                var dictionary = new Dictionary<int, int>();
                foreach (var clientId in clientIds)
                    dictionary[clientId] = ordersByInterval.Count(order => order.ClientId == clientId);

                var id = dictionary.Where(dict => dict.Value == dictionary.Values.Max()).Select(dict => dict.Key).FirstOrDefault();
                goldClient = EntityStorage.Instance.ClientStorage.FirstOrDefault(cl => cl.Id == id);
                if (goldClient == null)
                    throw new Exception("Золотой клиент не найден");
            }
            catch (Exception ex)
            {
                InnerException = ex;
            }
            return goldClient;
        }
    }
}
