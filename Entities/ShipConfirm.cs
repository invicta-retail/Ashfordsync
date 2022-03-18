using System;
using System.Collections.Generic;

namespace AshfordSync.Entities
{
    class ShipConfirm
    {
        public string orderNumber { get; set; }
        public string customerNumber { get; set; }
        public DateTime orderDate { get; set; }
        public List<Detail> details { get; set; }
    }
}
