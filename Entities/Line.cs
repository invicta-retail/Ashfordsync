namespace AshfordSync.Entities
{
    class Line
    {
        public int rmaLineId { get; set; }
        public string itemNumber { get; set; }
        public int quantity { get; set; }
        public string sourceOrderNumber { get; set; }
        public int sourceLineNumber { get; set; }
        public object smartLabelChargeCode { get; set; }
        public object smartLabelCharge { get; set; }
        public string reason { get; set; }
        public string restockCode { get; set; }
        public object comment { get; set; }
    }
}
