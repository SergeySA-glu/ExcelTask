using System;

namespace ExcelTask
{
    [ExcelSheet("Заявки")]
    public class Order : IExcelEntity
    {
        [ExcelHeader("Код заявки")]
        public int Id { get; set; }
        [ExcelHeader("Код товара")]
        public int WareId { get; set; }
        [ExcelHeader("Код клиента")]
        public int ClientId { get; set; }
        [ExcelHeader("Номер заявки")]
        public int Number { get; set; }
        [ExcelHeader("Требуемое количество")]
        public int Quantity { get; set; }
        [ExcelHeader("Дата размещения")]
        public DateTime Date { get; set; }

        #region IExcelEntity
        public string RowIndex { get; set; }
        #endregion
    }
}
