namespace BrickAtHeart.LUGTools.PicklistGenerator.Models
{
    public class Order
    {
        public Person Person { get; set; }

        public Part Part { get; set; }

        public int Quantity { get; set; }
    }
}