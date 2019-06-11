using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;

namespace BDO_Localisation_AddOn
{
    class TaxInvoice
    {
        private string su_field; //სერვისის მომხმარებელი
        private string sp_field; //სერვისის პაროლი
        private int user_id_field; //დეკლარირების მომხმარებლის უნიკალური ნომერი
        private int sua_field; //სერვისის მომხმარებლის უნიკალური ნომერი
        private TaxInvoiceService_HTTP.NtosService TaxInvoice_soapClient_field_HTTP = null;
        private TaxInvoiceService_HTTPS.NtosService TaxInvoice_soapClient_field_HTTPS = null;
        private string protocolType_field;
        private string tin_field;
        private int un_id;

        public TaxInvoice(string su, string sp, string protocolType)
        {
            if (protocolType == "HTTP")
            {
                this.TaxInvoice_soapClient_field_HTTP = new TaxInvoiceService_HTTP.NtosService();
            }
            else
            {
                this.TaxInvoice_soapClient_field_HTTPS = new TaxInvoiceService_HTTPS.NtosService();
            }
            this.su_field = su;
            this.sp_field = sp;
            this.protocolType_field = protocolType;
        }

        public TaxInvoice(string protocolType)
        {
            if (protocolType == "HTTP")
            {
                this.TaxInvoice_soapClient_field_HTTP = new TaxInvoiceService_HTTP.NtosService();
            }
            else
            {
                this.TaxInvoice_soapClient_field_HTTPS = new TaxInvoiceService_HTTPS.NtosService();
            }
            this.protocolType_field = protocolType;
        }

        public string su
        {
            get
            {
                return this.su_field;
            }
            set
            {
                this.su_field = value;
            }
        }

        public string sp
        {
            get
            {
                return this.sp_field;
            }
            set
            {
                this.sp_field = value;
            }
        }

        public int user_id
        {
            get
            {
                return this.user_id_field;
            }
            set
            {
                this.user_id_field = value;
            }
        }

        public int sua
        {
            get
            {
                return this.sua_field;
            }
            set
            {
                this.sua_field = value;
            }
        }

        public string tin
        {
            get
            {
                return this.tin_field;
            }
            set
            {
                this.tin_field = value;
            }
        }

        public TaxInvoiceService_HTTP.NtosService TaxInvoice_soapClient_HTTP
        {
            get
            {
                return this.TaxInvoice_soapClient_field_HTTP;
            }
            set
            {
                this.TaxInvoice_soapClient_field_HTTP = value;
            }
        }

        public TaxInvoiceService_HTTPS.NtosService TaxInvoice_soapClient_HTTPS
        {
            get
            {

                return this.TaxInvoice_soapClient_field_HTTPS;
            }
            set
            {
                this.TaxInvoice_soapClient_field_HTTPS = value;
            }
        }

        public string protocolType
        {
            get
            {
                return this.protocolType_field;
            }
            set
            {
                this.protocolType_field = value;
            }
        }

        /// <summary>მომხარებლის პაროლის შემოწმება</summary
        /// <param name="su_tmp">სერვისის მომხმარებელი</param>
        /// <param name="sp_tmp">სერვისის მომხმარებლის პაროლი</param>
        /// <param name="errorText"></param>
        /// <returns>user_id - დეკლარირების მომხმარებლის უნიკალური ნომერი, sua - სერვისის მომხმარებლის უნიკალური ნომერი</returns>
        public bool check_usr( string su_tmp, string sp_tmp, out string errorText)
        {
            errorText = null;
            int user_id_tmp = 0;
            int sua_tmp = 0;
            bool chek_service_user_tmp = false;

            try
            {
                if (protocolType == "HTTP")
                {
                    chek_service_user_tmp = TaxInvoice_soapClient_field_HTTP.chek(su_tmp, sp_tmp, ref user_id_tmp, out sua_tmp);
                }
                else
                {
                    chek_service_user_tmp = TaxInvoice_soapClient_field_HTTPS.chek(su_tmp, sp_tmp, ref user_id_tmp, out sua_tmp);
                }
            }
            catch (Exception ex)
            {
                errorText = ex.Message + "! chek_service_user()";
                return chek_service_user_tmp;
            }

            if (chek_service_user_tmp == true)
            {
                su = su_tmp;
                sp = sp_tmp;
                user_id = user_id_tmp;
                sua = sua_tmp;

                SAPbobsCOM.Recordset oRecordset = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                string query = "select \"FreeZoneNo\" from \"OADM\"";
                oRecordset.DoQuery(query);
                bool diplomat = false; 
                if (!oRecordset.EoF)
                {
                    tin = oRecordset.Fields.Item("FreeZoneNo").Value;
                    un_id = this.get_un_id_from_tin(null, out diplomat, out errorText);
                }
            }

            return chek_service_user_tmp;
        }

        //get_un_id_from_tin
        public int get_un_id_from_tin(string tinBP, out bool diplomat, out string errorText)
        {
            int un_id = 0;
            string name = "";
            bool isResult = false;
            string locTin = tinBP == null ? tin : tinBP;
            diplomat = false;

            errorText = null;
            try
            {
                if (protocolType == "HTTP")
                {
                    un_id = TaxInvoice_soapClient_field_HTTP.get_un_id_from_tin(user_id, locTin, su, sp, out name);
                }
                else
                {
                    un_id = TaxInvoice_soapClient_field_HTTPS.get_un_id_from_tin(user_id, locTin, su, sp, out name);
                }

                if (un_id > 0 || un_id < -3)
                {
                    isResult = true;
                }

                if (un_id < -15)
                {
                    diplomat = true;
                }

                if (isResult == false)
                {
                    un_id = 0;
                }
                return un_id;
            }
            catch (Exception ex)
            {
                errorText = ex.Message;
                return 0;
            }
        }

        //get_buyer_invoices
        /// <summary>
        /// 
        /// </summary>
        /// <param name="user_id"></param>
        /// <param name="un_id"></param>
        /// <param name="s_dt"></param>
        /// <param name="e_dt"></param>
        /// <param name="op_s_dt"></param>
        /// <param name="op_e_dt"></param>
        /// <param name="invoice_no"></param>
        /// <param name="sa_ident_no"></param>
        /// <param name="desc"></param>
        /// <param name="doc_mos_nom"></param>
        /// <param name="errorText"></param>
        /// <returns></returns>
        public DataTable get_buyer_invoices(DateTime s_dt, DateTime e_dt, DateTime op_s_dt, DateTime op_e_dt, string invoice_no, string sa_ident_no, string desc, string doc_mos_nom, out string errorText)
        {
            errorText = null;

            DataTable get_invoices_result = null;

            try
            {
                if (protocolType == "HTTP")
                {
                    get_invoices_result = TaxInvoice_soapClient_field_HTTP.get_buyer_invoices(user_id, un_id, s_dt, e_dt, op_s_dt, op_e_dt, invoice_no, sa_ident_no, desc, doc_mos_nom, su, sp);
                }
                else
                {
                    get_invoices_result = TaxInvoice_soapClient_field_HTTPS.get_buyer_invoices(user_id, un_id, s_dt, e_dt, op_s_dt, op_e_dt, invoice_no, sa_ident_no, desc, doc_mos_nom, su, sp);
                }
            }
            catch (Exception ex)
            {
                errorText = ex.Message + "! get_buyer_invoices()";
            }

            return get_invoices_result;
        }

        //get_seller_invoices
        /// <summary>
        /// 
        /// </summary>
        /// <param name="user_id"></param>
        /// <param name="un_id"></param>
        /// <param name="s_dt"></param>
        /// <param name="e_dt"></param>
        /// <param name="op_s_dt"></param>
        /// <param name="op_e_dt"></param>
        /// <param name="invoice_no"></param>
        /// <param name="sa_ident_no"></param>
        /// <param name="desc"></param>
        /// <param name="doc_mos_nom"></param>
        /// <param name="errorText"></param>
        /// <returns></returns>
        public DataTable get_seller_invoices(DateTime s_dt, DateTime e_dt, DateTime op_s_dt, DateTime op_e_dt, string invoice_no, string sa_ident_no, string desc, string doc_mos_nom, out string errorText)
        {
            errorText = null;

            System.Data.DataTable get_invoices_result = null;

            try
            {
                if (protocolType == "HTTP")
                {
                    get_invoices_result = TaxInvoice_soapClient_field_HTTP.get_seller_invoices(user_id, un_id, s_dt, e_dt, op_s_dt, op_e_dt, invoice_no, sa_ident_no, desc, doc_mos_nom, su, sp);
                }
                else
                {
                    get_invoices_result = TaxInvoice_soapClient_field_HTTPS.get_seller_invoices(user_id, un_id, s_dt, e_dt, op_s_dt, op_e_dt, invoice_no, sa_ident_no, desc, doc_mos_nom, su, sp);
                }
            }
            catch (Exception ex)
            {
                errorText = ex.Message + "! get_seller_invoices()";
            }

            return get_invoices_result;
        }

        //save_invoice
        /// <summary>
        /// 
        /// </summary>
        /// <param name="user_id"></param>
        /// <param name="invoice_id"></param>
        /// <param name="operation_date"></param>
        /// <param name="seller_un_id"></param>
        /// <param name="buyer_un_id"></param>
        /// <param name="overhead_no"></param>
        /// <param name="b_s_user_id"></param>
        /// <param name="errorText"></param>
        /// <returns></returns>
        public bool save_invoice(ref int invoice_id, DateTime operation_date, int buyer_un_id, string overhead_no, int b_s_user_id, out string errorText)
        {
            errorText = null;

            try
            {
                bool response = false;

                if (protocolType == "HTTP")
                {
                    response = TaxInvoice_soapClient_field_HTTP.save_invoice(user_id, ref invoice_id, operation_date, un_id, buyer_un_id, overhead_no, operation_date, b_s_user_id, su, sp);
                }
                else
                {
                    response = TaxInvoice_soapClient_field_HTTPS.save_invoice(user_id, ref invoice_id, operation_date, un_id, buyer_un_id, overhead_no, operation_date, b_s_user_id, su, sp);
                }

                if (response != true)
                {
                    //errorText = "ვერ მოხერხდა ანგარიშ-ფაქტურის შექმნა. შესაძლოა მყიდველი არ არის დღგ-ის გადამხდელი ან სერვისის მომხმარებელს არ აქვს საკმარისი უფლება ფაქტურის შექმნისთვის";
                    return response;
                }
                return response;
            }
            catch (Exception ex)
            {
                errorText = ex.Message;
                return false;
            }
        }

        public bool save_invoice_a(ref int invoice_id, DateTime operation_date, int buyer_un_id, string overhead_no, int b_s_user_id, out string errorText)
        {
            errorText = null;

            try
            {
                bool response = false;

                if (protocolType == "HTTP")
                {
                    response = TaxInvoice_soapClient_field_HTTP.save_invoice_a(user_id, ref invoice_id, operation_date, un_id, buyer_un_id, overhead_no, operation_date, b_s_user_id, su, sp);
                }
                else
                {
                    response = TaxInvoice_soapClient_field_HTTPS.save_invoice_a(user_id, ref invoice_id, operation_date, un_id, buyer_un_id, overhead_no, operation_date, b_s_user_id, su, sp);
                }

                if (response != true)
                {
                    //errorText = "ვერ მოხერხდა ანგარიშ-ფაქტურის შექმნა. შესაძლოა მყიდველი არ არის დღგ-ის გადამხდელი ან სერვისის მომხმარებელს არ აქვს საკმარისი უფლება ფაქტურის შექმნისთვის";
                    return response;
                }
                return response;
            }
            catch (Exception ex)
            {
                errorText = ex.Message;
                return false;
            }
        }


        //get_ntos_invoices_inv_nos
        /// <summary>
        /// ფაქტურის ზედნადებების მიღება
        /// </summary>
        /// <param name="invois_id"></param>
        /// <returns></returns>
        public DataTable get_ntos_invoices_inv_nos(int invois_id, out string errorText)
        {
            errorText = null;

            try
            {
                if (protocolType == "HTTP")
                {
                    return TaxInvoice_soapClient_field_HTTP.get_ntos_invoices_inv_nos(user_id, invois_id, su, sp);
                }
                else
                {
                    return TaxInvoice_soapClient_field_HTTPS.get_ntos_invoices_inv_nos(user_id, invois_id, su, sp);
                }
            }
            catch (Exception ex)
            {
                errorText = ex.Message;
                return null;
            }
        }

        //get_ntos_invoices_inv_nos
        /// <summary>
        /// ფაქტურის ზედნადებების მიღება
        /// </summary>
        /// <param name="invois_id"></param>
        /// <returns></returns>
        public DataTable get_ntos_invoices_inv_nos(int invois_id)
        {
            try
            {
                if (protocolType == "HTTP")
                {
                    return TaxInvoice_soapClient_field_HTTP.get_ntos_invoices_inv_nos(user_id, invois_id, su, sp);
                }
                else
                {
                    return TaxInvoice_soapClient_field_HTTPS.get_ntos_invoices_inv_nos(user_id, invois_id, su, sp);
                }
            }
            catch
            {            
                return null;
            }
        }

        //get_invoice
        /// <summary>
        /// 
        /// </summary>
        /// <param name="user_id"></param>
        /// <param name="invois_id"></param>
        /// <param name="errorText"></param>
        /// <returns></returns>
        public Dictionary<string, object> get_invoice(int invois_id, out string errorText)
        {
            errorText = null;

            Dictionary<string, object> ResponseStructure = null;

            string f_series = null;
            int f_number = -1;
            DateTime operation_dt = new DateTime(1, 1, 1);
            DateTime reg_dt = new DateTime(1, 1, 1);
            int seller_un_id = -1;
            int buyer_un_id = -1;
            string overhead_no = null;
            DateTime overhead_dt = new DateTime(1, 1, 1);
            int status = -1;
            string seq_num_s = null;
            string seq_num_b = null;
            int k_id = -1;
            int r_un_id = -1;
            int k_type = -1;
            int b_s_user_id = -1;
            int dec_status = -1;

            bool success = false;

            try
            {
                if (protocolType == "HTTP")
                {
                    success = TaxInvoice_soapClient_field_HTTP.get_invoice(user_id, invois_id, su, sp, out f_series, out f_number, out  operation_dt, out reg_dt, out seller_un_id, out buyer_un_id, out overhead_no, out
                                                   overhead_dt, out status, out seq_num_s, out seq_num_b, out k_id, out r_un_id, out k_type, out b_s_user_id, out dec_status);
                }
                else
                {
                    success = TaxInvoice_soapClient_field_HTTPS.get_invoice(user_id, invois_id, su, sp, out f_series, out f_number, out  operation_dt, out reg_dt, out seller_un_id, out buyer_un_id, out overhead_no, out
                                                   overhead_dt, out status, out seq_num_s, out seq_num_b, out k_id, out r_un_id, out k_type, out b_s_user_id, out dec_status);
                }

                ResponseStructure = new Dictionary<string, object>();

                if (f_series == null || success == false)
                {
                    ResponseStructure.Add("reg_dt", new DateTime(1, 1, 1));
                    ResponseStructure.Add("f_number", -1);
                    ResponseStructure.Add("f_series", null);
                    ResponseStructure.Add("result", true);
                    ResponseStructure.Add("status", -1);
                    ResponseStructure.Add("seq_num_b", null);
                    ResponseStructure.Add("seq_num_s", null);
                    ResponseStructure.Add("operation_dt", new DateTime(1, 1, 1));
                    ResponseStructure.Add("seller_un_id", -1);
                    ResponseStructure.Add("buyer_un_id", -1);
                    ResponseStructure.Add("overhead_no", null);
                    ResponseStructure.Add("k_id", -1);
                    ResponseStructure.Add("r_un_id", -1);
                    ResponseStructure.Add("k_type", -1);
                    ResponseStructure.Add("b_s_user_id", -1);
                    ResponseStructure.Add("dec_status", -1);                 
                }
                else
                {
                    ResponseStructure.Add("reg_dt", reg_dt);
                    ResponseStructure.Add("f_number", f_number);
                    ResponseStructure.Add("f_series", f_series);
                    ResponseStructure.Add("result", true);
                    ResponseStructure.Add("status", status);
                    ResponseStructure.Add("seq_num_b", seq_num_b);
                    ResponseStructure.Add("seq_num_s", seq_num_s);
                    ResponseStructure.Add("operation_dt", operation_dt);
                    ResponseStructure.Add("seller_un_id", seller_un_id);
                    ResponseStructure.Add("buyer_un_id", buyer_un_id);
                    ResponseStructure.Add("overhead_no", overhead_no);
                    ResponseStructure.Add("k_id", k_id);
                    ResponseStructure.Add("r_un_id", r_un_id);
                    ResponseStructure.Add("k_type", k_type);
                    ResponseStructure.Add("b_s_user_id", b_s_user_id);
                    ResponseStructure.Add("dec_status", dec_status);
                }
                return ResponseStructure;
            }
            catch (Exception ex)
            {
                errorText = ex.Message;
                return ResponseStructure;
            }           
        }

        //get_seq_nums
        /// <summary>
        /// get_seq_nums თვის მიხედვით დეკლარაციის ნომრის დაბრუნება
        /// </summary>
        /// <param name="period"></param>
        /// <returns></returns>
        public DataTable get_seq_nums(string period, out string errorText)
        {
            errorText = null;

            if (protocolType == "HTTP")
            {
                return TaxInvoice_soapClient_field_HTTP.get_seq_nums(period, user_id, su, sp);
            }
            else
            {
                return TaxInvoice_soapClient_field_HTTPS.get_seq_nums(period, user_id, su, sp);
            }
        }

        //add_inv_to_decl
        /// <summary>
        /// დეკლარაციაზე დამატება add_inv_to_decl
        /// </summary>
        /// <param name="seq_num"></param>
        /// <param name="inv_id"></param>
        /// <param name="errorText"></param>
        /// <returns></returns>
        public bool add_inv_to_decl(int seq_num, int inv_id, out string errorText)
        {
            errorText = null;

            if (protocolType == "HTTP")
            {
                return TaxInvoice_soapClient_field_HTTP.add_inv_to_decl(user_id, seq_num, inv_id, su, sp);
            }
            else
            {
                return TaxInvoice_soapClient_field_HTTPS.add_inv_to_decl(user_id, seq_num, inv_id, su, sp);
            }
        }

        //change_invoice_status
        /// <summary>
        /// change_invoice_status
        /// </summary>
        public bool change_invoice_status(int inv_id, int status, out string errorText)
        {
            errorText = null;

            if (protocolType == "HTTP")
            {
                return TaxInvoice_soapClient_field_HTTP.change_invoice_status(user_id, inv_id, status, su, sp);
            }
            else
            {
                return TaxInvoice_soapClient_field_HTTPS.change_invoice_status(user_id, inv_id, status, su, sp);
            }
        }

        //acsept_invoice_status
        /// <summary>
        /// ფაქტურის დადასტურება
        /// </summary>
        /// <param name="errorText"></param>
        public bool acsept_invoice_status(int inv_id, int status, out string errorText)
        {
            errorText = null;

            if (protocolType == "HTTP")
            {
                return TaxInvoice_soapClient_field_HTTP.acsept_invoice_status(user_id, inv_id, status, su, sp);
            }
            else
            {
                return TaxInvoice_soapClient_field_HTTPS.acsept_invoice_status(user_id, inv_id, status, su, sp);
            }
        }

        //acsept_invoice_request_status
        /// <summary>
        /// acsept_invoice_request_status
        /// </summary>
        /// <param name="errorText"></param>
        public bool acsept_invoice_request_status(int id, int seller_un_id, out string errorText)
        {
            errorText = null;

            if (protocolType == "HTTP")
            {
                return TaxInvoice_soapClient_field_HTTP.acsept_invoice_request_status(id, user_id, un_id, su, sp);
            }
            else
            {
                return TaxInvoice_soapClient_field_HTTPS.acsept_invoice_request_status(id, user_id, un_id, su, sp);
            }
        }

        //ref_invoice_status
        /// <summary>
        /// ref_invoice_status უარყოფა
        /// </summary>
        public bool ref_invoice_status(int inv_id, string ref_text, out string errorText)
        {
            errorText = null;

            if (protocolType == "HTTP")
            {
                return TaxInvoice_soapClient_field_HTTP.ref_invoice_status(user_id, inv_id, ref_text, su, sp);
            }
            else
            {
                return TaxInvoice_soapClient_field_HTTPS.ref_invoice_status(user_id, inv_id, ref_text, su, sp);
            }
        }

        //k_invoice
        /// <summary>
        /// კორექტირება
        /// </summary>
        /// <param name="errorText"></param>
        public bool k_invoice(int inv_id, int k_type, out int k_id, out string errorText)
        {
            errorText = null;

            if (protocolType == "HTTP")
            {
                return TaxInvoice_soapClient_field_HTTP.k_invoice(user_id, inv_id, k_type, su, sp, out k_id);
            }
            else
            {
                return TaxInvoice_soapClient_field_HTTPS.k_invoice(user_id, inv_id, k_type, su, sp, out k_id);
            }
        }


        //save_ntos_invoices_inv_nos
        /// <summary>
        /// ზედნადების დამატება
        /// </summary>
        public bool save_ntos_invoices_inv_nos(int invois_id, string overhead_no, DateTime overhead_dt, out string errorText)
        {
            errorText = null;

            try
            {
                if (protocolType == "HTTP")
                {
                    return TaxInvoice_soapClient_field_HTTP.save_ntos_invoices_inv_nos(invois_id, user_id, overhead_no, overhead_dt, su, sp);
                }
                else
                {
                    return TaxInvoice_soapClient_field_HTTPS.save_ntos_invoices_inv_nos(invois_id, user_id, overhead_no, overhead_dt, su, sp);
                }
            }
            catch (Exception ex)
            {
                errorText = ex.Message;
                return false;
            }
        }

        //delete_ntos_invoices_inv_nos
        /// <summary>
        /// ზედნადების წაშლა
        /// </summary>
        public bool delete_ntos_invoices_inv_nos(int id, int inv_id, out string errorText)
        {
            errorText = null;
            try
            {
                if (protocolType == "HTTP")
                {
                    return TaxInvoice_soapClient_field_HTTP.delete_ntos_invoices_inv_nos(user_id, id, inv_id, su, sp);
                }
                else
                {
                    return TaxInvoice_soapClient_field_HTTPS.delete_ntos_invoices_inv_nos(user_id, id, inv_id, su, sp);
                }
            }
            catch (Exception ex)
            {
                errorText = ex.Message;
                return false;
            }
        }

        //save_invoice_desc
        /// <summary>
        /// სტრიქონის ფაქტურაში დამატება
        /// </summary>
        /// <param name="errorText"></param>
        public bool save_invoice_desc(int id, int invois_id, string goods, string g_unit, decimal g_number, decimal full_amount, decimal drg_amount, decimal aqcizi_amount, int akciz_id, out string errorText)
        {
            errorText = null;
            try
            {
                if (protocolType == "HTTP")
                {
                    return TaxInvoice_soapClient_field_HTTP.save_invoice_desc(user_id, ref id, su, sp, invois_id, goods, g_unit, g_number, full_amount, drg_amount, aqcizi_amount, akciz_id);
                }
                else
                {
                    return TaxInvoice_soapClient_field_HTTPS.save_invoice_desc(user_id, ref id, su, sp, invois_id, goods, g_unit, g_number, full_amount, drg_amount, aqcizi_amount, akciz_id);
                }
            }
            catch (Exception ex)
            {
                errorText = ex.Message;
                return false;
            }
        }

        //delete_invoice_desc
        /// <summary>
        /// სტრიქონის წაშლა ფაქტურიდან 
        /// </summary>
        /// <param name="errorText"></param>
        /// <returns></returns>
        public bool delete_invoice_desc(int id, int inv_id, out string errorText)
        {
            errorText = null;
            try
            {
                if (protocolType == "HTTP")
                {
                    return TaxInvoice_soapClient_field_HTTP.delete_invoice_desc(user_id, id, inv_id, su, sp);
                }
                else
                {
                    return TaxInvoice_soapClient_field_HTTPS.delete_invoice_desc(user_id, id, inv_id, su, sp);
                }
            }
            catch (Exception ex)
            {
                errorText = ex.Message;
                return false;
            }
        }

        //get_ser_users_notes
        /// <summary>
        /// სერვისი მომხმარებლების მიღება
        /// </summary>
        public DataTable get_ser_users_notes(out string errorText)
        {
            errorText = null;

            if (protocolType == "HTTP")
            {
                return TaxInvoice_soapClient_field_HTTP.get_ser_users_notes(tin);
            }
            else
            {
                return TaxInvoice_soapClient_field_HTTPS.get_ser_users_notes(tin);
            }
        }

        //get_invoice_desc
        /// <summary>
        /// ფაქტურის ცხრილური ნაწილის მიღება
        /// </summary>
        /// <param name="errorText"></param>
        public DataTable get_invoice_desc(int invois_id, out string errorText)
        {
            errorText = null;
            try
            {
                if (protocolType == "HTTP")
                {
                    return TaxInvoice_soapClient_field_HTTP.get_invoice_desc(user_id, invois_id, su, sp);
                }
                else
                {
                    return TaxInvoice_soapClient_field_HTTPS.get_invoice_desc(user_id, invois_id, su, sp);
                }
            }
            catch (Exception ex)
            {
                errorText = ex.Message;  
                return null;
            }
        }

        //get_invoice_desc
        /// <summary>
        /// ფაქტურის ცხრილური ნაწილის მიღება
        /// </summary>
        public DataTable get_invoice_desc(int invois_id)
        {     
            try
            {
                if (protocolType == "HTTP")
                {
                    return TaxInvoice_soapClient_field_HTTP.get_invoice_desc(user_id, invois_id, su, sp);
                }
                else
                {
                    return TaxInvoice_soapClient_field_HTTPS.get_invoice_desc(user_id, invois_id, su, sp);
                }
            }
            catch
            {          
                return null;
            }
        }

        //get_requested_invoices
        /// <summary>
        /// ფაქტურის გამოწერაზე მოთხოვნის მიღება
        /// </summary>
        /// <param name="errorText"></param>
        public DataTable get_requested_invoices(out string errorText)
        {
            errorText = null;

            if (protocolType == "HTTP")
            {
                return TaxInvoice_soapClient_field_HTTP.get_requested_invoices(user_id, un_id, su, sp);
            }
            else
            {
                return TaxInvoice_soapClient_field_HTTPS.get_requested_invoices(user_id, un_id, su, sp);
            }

        }

        //get_makoreqtirebeli
        /// <summary>
        /// მაკორექტირებელი ფაქტურის მიღება
        /// </summary>
        /// <param name="errorText"></param>
        public bool get_makoreqtirebeli(int inv_id, out int k_id, out string errorText)
        {
            errorText = null;

            if (protocolType == "HTTP")
            {
                return TaxInvoice_soapClient_field_HTTP.get_makoreqtirebeli(user_id, inv_id, su, sp, out k_id);
            }
            else
            {
                return TaxInvoice_soapClient_field_HTTPS.get_makoreqtirebeli(user_id, inv_id, su, sp, out k_id);
            }
        }

    }
}
