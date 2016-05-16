namespace BOM
{
    public class Child
    {
        public string ItemCode { get; set; }
        public string ItemDesc { get; set; }
        public string Quantity { get; set; }
        public string Cost { get; set; }
        public Child(string ItemCode,string ItemDesc, string Quantity,string Cost)
        {
            this.ItemCode = ItemCode;
            this.ItemDesc = ItemDesc;
            this.Quantity = Quantity;
            this.Cost = Cost;
        }
    }
}