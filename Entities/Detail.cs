using System;
using System.Collections.Generic;

namespace AshfordSync.Entities
{
    class Detail
    {
        public int lineNumber { get; set; }
        public string itemNumber { get; set; }
        public int orderedQuantity { get; set; }
        public int shippedQuantity { get; set; }
        public int canceledQuantity { get; set; }
        public DateTime shippedDate { get; set; }
        public string carrier { get; set; }
        public string trackingNumber { get; set; }
        public bool prePaidReturnLabelUsed { get; set; }
        public Decimal prePaidReturnLabelCost { get; set; }

        public static implicit operator List<object>(Detail v)
        {
            throw new NotImplementedException();
        }
    }
}
