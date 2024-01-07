namespace ExcelTask
{
    [ExcelSheet("Товары")]
    public class Ware : IExcelEntity
    {
        [ExcelHeader("Код товара")]
        public int Id { get; set; }
        [ExcelHeader("Наименование")]
        public string Name { get; set; }
        [ExcelHeader("Ед. измерения")]
        public string Units { get; set; }
        [ExcelHeader("Цена товара за единицу")]
        public decimal Price { get; set; }

        #region IExcelEntity
        public string RowIndex { get; set; }
        #endregion
    }
}
