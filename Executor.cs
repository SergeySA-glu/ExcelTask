using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;

namespace ExcelTask
{
    public class Executor
    {
        private static Executor _instance;
        public static Executor Instance => _instance ?? (_instance = new Executor());

        private bool _dataLoaded;

        public Exception InnerException { get; set; }

        public bool LoadEntities(string pathFile)
        {
            try
            {
                if (_dataLoaded && EntityStorage.Instance.PathFile == pathFile)
                    throw new Exception("Данные уже были загружены");

                if (_dataLoaded)
                {
                    EntityStorage.Instance.ClearStorage();
                    _dataLoaded = false;
                }

                using (var document = SpreadsheetDocument.Open(pathFile, false))
                {
                    var workbookPart = document.WorkbookPart;
                    var workbook = workbookPart.Workbook;
                    var sheets = workbook.Elements<Sheets>().First();
                    foreach (var type in Assembly.GetExecutingAssembly().GetTypes().Where(t => t.GetInterfaces().Contains(typeof(IExcelEntity))))
                    {
                        var sheet = sheets.Elements<Sheet>().SingleOrDefault(s => s.Name == ExcelHelper.GetExcelSheetNameByType(type));
                        var workSheetPart = (WorksheetPart)document.WorkbookPart.GetPartById(sheet.Id.Value);
                        var sheetData = workSheetPart.Worksheet.Elements<SheetData>().First();

                        foreach (var row in sheetData.Elements<Row>())
                        {
                            if (row.Equals(sheetData.Elements<Row>().First()))
                            {
                                EntityStorage.Instance.SetSheetHeaders(row, workbookPart);
                            }
                            else
                            {
                                var entity = Activator.CreateInstance(type) as IExcelEntity;
                                if (entity != null)
                                    EntityStorage.Instance.SaveRowInStorage(entity, type, row, workbookPart);
                            }
                        }
                    }
                }
                _dataLoaded = true;
                EntityStorage.Instance.PathFile = pathFile;
            }
            catch (Exception ex)
            {
                InnerException = ex;
                return false;
            }
            return _dataLoaded;
        }

        public List<(DateTime, Client, int, decimal)> GetClientsByWare(string wareName)
        {
            var ordersInfo = new List<(DateTime, Client, int, decimal)>();
            try
            {
                if (!_dataLoaded)
                    throw new Exception("Данные не загружены. Требуется выполнить запрос #1");

                var ware = EntityStorage.Instance.TypeStorages[typeof(Ware)].Cast<Ware>().FirstOrDefault(w => w.Name == wareName.Trim());
                if (ware == null)
                    throw new Exception("Товар не найден");

                var orders = EntityStorage.Instance.TypeStorages[typeof(Order)].Cast<Order>().Where(order => order.WareId == ware.Id);
                if (!orders.Any())
                    throw new Exception("Заказов по указанному товару не найдено");

                foreach (var order in orders)
                {
                    ordersInfo.Add(
                        (order.Date,
                        EntityStorage.Instance.TypeStorages[typeof(Client)].Cast<Client>().FirstOrDefault(cl => cl.Id == order.ClientId),
                        order.Quantity,
                        ware.Price));
                }
            }
            catch (Exception ex)
            {
                InnerException = ex;
            }
            return ordersInfo;
        }

        public bool SetNewContact(string organization, string newContact)
        {
            try
            {
                if (!_dataLoaded)
                    throw new Exception("Данные не загружены. Требуется выполнить запрос #1");

                var client = EntityStorage.Instance.TypeStorages[typeof(Client)].Cast<Client>().FirstOrDefault(cl => cl.Name == organization.Trim());
                if (client == null)
                    throw new Exception("Клиент не найден");

                var columnName = EntityStorage.Instance.ColumnCharTree[typeof(Client)][typeof(Client).GetProperty("Contact")];
                var rowIndex = client.RowIndex;

                using (var document = SpreadsheetDocument.Open(EntityStorage.Instance.PathFile, true))
                {
                    var workbookPart = document.WorkbookPart;
                    var workbook = workbookPart.Workbook;
                    var sheets = workbook.Elements<Sheets>().First();
                    var sheet = sheets.Elements<Sheet>().SingleOrDefault(s => s.Name == ExcelHelper.GetExcelSheetNameByType(typeof(Client)));
                    var workSheetPart = (WorksheetPart)document.WorkbookPart.GetPartById(sheet.Id.Value);
                    var sheetData = workSheetPart.Worksheet.Elements<SheetData>().First();

                    var row = sheetData.Elements<Row>().SingleOrDefault(r => r.RowIndex == rowIndex);
                    var cell = row.Elements<Cell>().SingleOrDefault(c => c.CellReference == columnName + rowIndex);

                    if (cell.DataType.Value == CellValues.SharedString)
                    {
                        var sharedStringTable = workbookPart.SharedStringTablePart.SharedStringTable;
                        var index = 0;
                        foreach (var sharedStringItem in sharedStringTable.Elements<SharedStringItem>())
                        {
                            if (index == int.Parse(cell.CellValue.Text))
                                sharedStringItem.Text = new Text(newContact);
                            index++;
                        }
                    }
                    else
                    {
                        cell.CellValue = new CellValue(newContact);
                    }
                    document.WorkbookPart.Workbook.Save();
                }

                client.Contact = newContact;
                return true;
            }
            catch (Exception ex)
            {
                InnerException = ex;
                return false;
            }
        }

        public Client GetGoldClient(bool year, int dateValue)
        {
            Client goldClient = null;
            try
            {
                if (!_dataLoaded)
                    throw new Exception("Данные не загружены. Требуется выполнить запрос #1");

                var ordersByInterval = EntityStorage.Instance.TypeStorages[typeof(Order)].Cast<Order>()
                    .Where(order => year ?
                        order.Date.Year == dateValue :
                        order.Date.Month == dateValue && order.Date != default);

                var clientIds = ordersByInterval.Select(order => order.ClientId).Distinct();

                var dictionary = new Dictionary<int, int>();
                foreach (var clientId in clientIds)
                    dictionary[clientId] = ordersByInterval.Count(order => order.ClientId == clientId);

                var id = dictionary.Where(dict => dict.Value == dictionary.Values.Max()).Select(dict => dict.Key).FirstOrDefault();
                goldClient = EntityStorage.Instance.TypeStorages[typeof(Client)].Cast<Client>().FirstOrDefault(cl => cl.Id == id);
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
