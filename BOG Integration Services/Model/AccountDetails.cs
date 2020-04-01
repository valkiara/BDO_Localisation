namespace BDO_Localisation_AddOn.BOG_Integration_Services.Model
{
    public class AccountDetails
    {
        private string _name;

        public string Name
        {
            get => _name.Replace("quot;",@"""").Replace("apos; apos;", "''").Replace("apos;", "'");
            set => _name = value;
        }

        public string Inn { get; set; }
        public string AccountNumber { get; set; }
        public string BankCode { get; set; }
        public string BankName { get; set; }

        public override string ToString()
        {
            return string.Format("{0} {1} {2} {3} {4}", Name, Inn, AccountNumber, BankCode, BankName);
        }
    }
}