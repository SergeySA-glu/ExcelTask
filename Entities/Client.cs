namespace ExcelTask
{
    [ExcelSheet("Клиенты")]
    public class Client : IExcelEntity
    {
        [ExcelHeader("Код клиента")]
        public int Id { get; set; }
        [ExcelHeader("Наименование организации")]
        public string Name { get; set; }
        [ExcelHeader("Адрес")]
        public string Address { get; set; }
        [ExcelHeader("Контактное лицо (ФИО)")]
        public string Contact { get; set; }

        #region IExcelEntity
        public string RowIndex { get; set; }
        #endregion
    }
}
