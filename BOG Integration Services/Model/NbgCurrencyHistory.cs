using System;

namespace BDO_Localisation_AddOn.BOG_Integration_Services.Model
{
    public class NbgCurrencyHistory
    {
        public DateTime Date { get; set; }
        public decimal Rate { get; set; }
        public string Currency { get; set; }
    }
}