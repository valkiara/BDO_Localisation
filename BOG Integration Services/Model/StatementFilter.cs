using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BDO_Localisation_AddOn.BOG_Integration_Services.Model
{
    public class StatementFilter
    {
        public string AccountNumber { get; set; }
        public string Currency { get; set; }
        public DateTime PeriodFrom { get; set; }
        public DateTime PeriodTo { get; set; }
        public int Page { get; set; }
    }
}
