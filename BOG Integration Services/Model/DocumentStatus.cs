using System;
using System.Collections.Generic;

namespace BDO_Localisation_AddOn.BOG_Integration_Services.Model
{
    public class DocumentStatus
    {
        public Guid UniqueId { get; set; }

        public long? UniqueKey { get; set; }

        public string Status { get; set; }

        public string BulkLineStatus { get; set; }

        public int? RejectCode { get; set; }

        public int? ResultCode { get; set; }
    }

    public class BulkPaymentStatus
    {
        public string Status { get; set; }

        public List<DocumentStatus> DocumentStatuses { get; set; }
    }
}