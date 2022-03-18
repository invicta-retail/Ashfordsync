namespace AshfordSync.Entities
{
    class Parameters
    {
        public int initialRow { get; set; }
        public int invItemNameColumn { get; set; }
        public int invQuantityColumn { get; set; }
        public string tcouri { get; set; }
        public int supplierId { get; set; }
        public string defaultcarrier { get; set; }
        public string connectionString { get; set; }
        public string validDomain { get; set; }
        public int scOrderNumberColumn { get; set; }
        public int scCustomerNumberColumn { get; set; }
        public int scOrderDateColumn { get; set; }
        public int scLineNumberColumn { get; set; }
        public int scItemNumberColumn { get; set; }
        public int scOrderedQuantityColumn { get; set; }
        public int scShippedQuantityColumn { get; set; }
        public int scCancelledQuantityColumn { get; set; }
        public int scShippedDateColumn { get; set; }
        public int scCarrierColumn { get; set; }
        public int scTrackingNumberColumn { get; set; }
        public int scPrePaidRetunLabelUsedColumn { get; set; }
        public int scPrePaidReturnLabelCostColumn { get; set; }
        public int rmaHeaderIdColumn { get; set; }
        public int rmaOrderNumberColumn { get; set; }
        public int rmaDateColumn { get; set; }
        public int rmaRetailingEnterprisesRmaNumberColumn { get; set; }
        public int rmaLineIdColumn { get; set; }
        public int rmaItemNumberColumn { get; set; }
        public int rmaQuantityColumn { get; set; }
        public int rmaSourceOrderNumberColumn { get; set; }
        public int rmaSourceLineNumberColumn { get; set; }
        public int rmaSmartLabelChargeCodeColumn { get; set; }
        public int rmaSmartLabelChargeColumn { get; set; }
        public int rmaReasonColumn { get; set; }
        public int rmaRestockCodeColumn { get; set; }
        public int rmaCommentColumn { get; set; }

        public string spaceReplacement { get; set; }
    }
}
