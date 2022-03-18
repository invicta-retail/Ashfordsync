using System;
using System.Collections.Generic;

namespace AshfordSync.Entities
{
    class RMA
    {
        public int rmaHeaderId { get; set; }
        public string rmaOrderNumber { get; set; }
        public DateTime rmaDate { get; set; }
        public object retailingEnterprisesRmaNumber { get; set; }
        public List<Line> lines { get; set; }
    }
}
