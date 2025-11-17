using System.Collections.Generic;

namespace DEFOR_Combinations
{
    public class CombinationModel
    {
        public string Name { get; set; }
        public List<CombinationItem> Items { get; set; }

        public CombinationModel()
        {
            Items = new List<CombinationItem>();
        }

        public override string ToString()
        {
            return Name;
        }
    }

    public class CombinationItem
    {
        public string FamilyName { get; set; }
        public string SymbolName { get; set; }
        public string SymbolUniqueId { get; set; }
        public int Quantity { get; set; }
        public double OffsetX { get; set; }
        public double OffsetY { get; set; }

        public override string ToString()
        {
            return $"{FamilyName} : {SymbolName} (Qty: {Quantity}, dX: {OffsetX}, dY: {OffsetY})";
        }
    }
}
