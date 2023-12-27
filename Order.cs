using System;

namespace ExcelTask
{
    public class Order
    {
        public int Id { get; set; }
        public int WareId { get; set; }
        public int ClientId { get; set; }
        public int Number { get; set; }
        public int Quantity { get; set; }
        public DateTime Date { get; set; }
    }
}
