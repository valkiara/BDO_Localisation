using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Xml;
using System.Xml.Linq;
using System.Data;
using System.IO;

namespace BDO_Localisation_AddOn
{
    class WayBill
    {
        //private string user_name_field; //ელექტრონული დეკლარირების მომხმარებლის სახელი
        //private string user_password_field; //ელექტრონული დეკლარირების მომხმარებლის პაროლი
        private string su_field; //სერვისის მომხმარებელი
        private string sp_field; //სერვისის პაროლი
        private int un_id_field; //გადამხდელის უნიკალური ნომერი
        private int user_id_field; //სერვისის მომხმარებლის ID
        private WayBillService_HTTP.WayBills wayBill_soapClient_field_HTTP = null;
        private WayBillService_HTTPS.WayBills wayBill_soapClient_field_HTTPS = null;
        private string protocolType_field;

        public WayBill(string su, string sp, string protocolType)
        {
            if (protocolType == "HTTP")
            {
                this.wayBill_soapClient_field_HTTP = new WayBillService_HTTP.WayBills();
            }
            else
            {
                this.wayBill_soapClient_field_HTTPS = new WayBillService_HTTPS.WayBills();
            }

            this.su_field = su;
            this.sp_field = sp;
            this.protocolType_field = protocolType;
        }

        public WayBill(string protocolType)
        {
            if (protocolType == "HTTP")
            {
                this.wayBill_soapClient_field_HTTP = new WayBillService_HTTP.WayBills();
            }
            else
            {
                this.wayBill_soapClient_field_HTTPS = new WayBillService_HTTPS.WayBills();
            }
            this.protocolType_field = protocolType;
        }

        //public string user_name
        //{
        //    get
        //    {
        //        return this.user_name_field;
        //    }
        //    set
        //    {
        //        this.user_name_field = value;
        //    }
        //}

        //public string user_password
        //{
        //    get
        //    {
        //        return this.user_password_field;
        //    }
        //    set
        //    {
        //        this.user_password_field = value;
        //    }
        //}

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

        public int un_id
        {
            get
            {
                return this.un_id_field;
            }
            set
            {
                this.un_id_field = value;
            }
        }

        public int un_user_id
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

        public WayBillService_HTTP.WayBills wayBill_soapClient_HTTP
        {
            get
            {

                return this.wayBill_soapClient_field_HTTP;
            }
            set
            {
                this.wayBill_soapClient_field_HTTP = value;
            }
        }

        public WayBillService_HTTPS.WayBills wayBill_soapClient_HTTPS
        {
            get
            {

                return this.wayBill_soapClient_field_HTTPS;
            }
            set
            {
                this.wayBill_soapClient_field_HTTPS = value;
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

        //სერვისის მომხმარებლის ადმინისტრირება --->

        /// <summary> მომხმარებლის შექმნა</summary>
        /// <param name="user_name">ელექტრონული დეკლარირების მომხმარებლის სახელი</param>
        /// <param name="user_password">ელექტრონული დეკლარირების მომხმარებლის პაროლი</param>
        /// <param name="ip_str">IP საიდანც მოხდება სერვისების გამოყენება</param>
        /// <param name="su_str">სერვისის მომხმარებელი</param> 
        /// <param name="sp_str">სერვისის მომხმარებლის პაროლი</param>
        /// <param name="name">ობიექტის სახელი</param>
        /// <param name="errorText"></param>
        /// <returns>ლოგიკურ ცვლადს - true მომხმარებელი შეიქმნა, false მომხმარფებელი არ შეიქმნა</returns>
        public bool create_service_user(string user_name, string user_password, string ip_str, string su_str, string sp_str, string name_str, out string errorText)
        {
            errorText = null;
            bool create_service_user_result = false;
            try
            {
                if (protocolType == "HTTP")
                {
                    create_service_user_result = wayBill_soapClient_HTTP.create_service_user(user_name, user_password, ip_str, name_str, su_str, sp_str);
                }
                else
                {
                    create_service_user_result = wayBill_soapClient_HTTPS.create_service_user(user_name, user_password, ip_str, name_str, su_str, sp_str);
                }
            }
            catch (Exception ex)
            {
                errorText = ex.Message + "! create_service_user()";
                return create_service_user_result;
            }

            return create_service_user_result;
        }

        /// <summary>მომხმარებლის განახლება</summary>
        /// <param name="user_name">ელექტრონული დეკლარირების მომხმარებლის სახელი</param>
        /// <param name="user_password">ელექტრონული დეკლარირების მომხმარებლის პაროლი</param>
        /// <param name="ip_str">IP საიდანც მოხდება სერვისების გამოყენება</param>
        /// <param name="su_str">სერვისის მომხმარებელი</param> 
        /// <param name="sp_str">სერვისის მომხმარებლის პაროლი</param>
        /// <param name="name">ობიექტის სახელი</param>
        /// <param name="errorText"></param>
        /// <returns> ლოგიკურ ცვლადს - true მომხმარებელის მონაცემები განახლდა, false არ განახლდა</returns>
        public bool update_service_user(string user_name, string user_password, string ip_str, string su_str, string sp_str, string name_str, out string errorText)
        {
            errorText = null;
            bool update_service_user_result = false;
            try
            {
                if (protocolType == "HTTP")
                {
                    update_service_user_result = wayBill_soapClient_HTTP.update_service_user(user_name, user_password, ip_str, name_str, su_str, sp_str);
                }
                else
                {
                    update_service_user_result = wayBill_soapClient_HTTPS.update_service_user(user_name, user_password, ip_str, name_str, su_str, sp_str);
                }
            }
            catch (Exception ex)
            {
                errorText = ex.Message + "! update_service_user()";
                return update_service_user_result;
            }

            return update_service_user_result;
        }

        /// <summary>მომხმარებლების სიის გამოტანა</summary>
        /// <param name="user_name">ელექტრონული დეკლარირების მომხმარებლის სახელი</param>
        /// <param name="user_password">ელექტრონული დეკლარირების მომხმარებლის პაროლი</param>
        /// <param name="errorText"></param>
        /// <returns>Dictionary, ან NULL</returns>
        public Dictionary<string, HashSet<string>> get_service_users(string user_name, string user_password, out string errorText)
        {
            errorText = null;
            Dictionary<string, HashSet<string>> service_users_map = null;
            XmlNode get_service_users_result = null;
            try
            {
                if (protocolType == "HTTP")
                {
                    get_service_users_result = wayBill_soapClient_HTTP.get_service_users(user_name, user_password);
                }
                else
                {
                    get_service_users_result = wayBill_soapClient_HTTPS.get_service_users(user_name, user_password);
                }
            }
            catch (Exception ex)
            {
                errorText = ex.Message + "! get_service_users()";
                return service_users_map;
            }

            if (get_service_users_result != null)
            {
                service_users_map = new Dictionary<string, HashSet<string>>();

                XmlNodeList itemNodes = get_service_users_result.SelectNodes("ServiceUser");
                foreach (XmlNode itemNode in itemNodes)
                {
                    string ID = (itemNode.SelectSingleNode("ID") == null) ? "" : itemNode.SelectSingleNode("ID").InnerText;
                    string USER_NAME = (itemNode.SelectSingleNode("USER_NAME") == null) ? "" : itemNode.SelectSingleNode("USER_NAME").InnerText;
                    string UN_ID = (itemNode.SelectSingleNode("UN_ID") == null) ? "" : itemNode.SelectSingleNode("UN_ID").InnerText;
                    string IP = (itemNode.SelectSingleNode("IP") == null) ? "" : itemNode.SelectSingleNode("IP").InnerText;
                    string NAME = (itemNode.SelectSingleNode("NAME") == null) ? "" : itemNode.SelectSingleNode("NAME").InnerText;
                    service_users_map.Add(ID, new HashSet<string>() { USER_NAME, UN_ID, IP, NAME });
                }
            }

            return service_users_map;
        }

        /// <summary>მომხარებლის პაროლის შემოწმება</summary
        /// <param name="su_tmp">სერვისის მომხმარებელი</param>
        /// <param name="sp_tmp">სერვისის მომხმარებლის პაროლი</param>
        /// <param name="errorText"></param>
        /// <returns>un_id - გადამხდელის უნიკალური ნომერი, s_user_id - სერვისის მომხმარებლის ID</returns>
        public bool chek_service_user(string su_tmp, string sp_tmp, out string errorText)
        {
            errorText = null;
            int un_id_tmp = 0;
            int un_user_id_tmp = 0;
            bool chek_service_user_tmp = false;

            try
            {
                if (protocolType == "HTTP")
                {
                    chek_service_user_tmp = wayBill_soapClient_HTTP.chek_service_user(su_tmp, sp_tmp, out un_id_tmp, out un_user_id_tmp);
                }
                else
                {
                    chek_service_user_tmp = wayBill_soapClient_HTTPS.chek_service_user(su_tmp, sp_tmp, out un_id_tmp, out un_user_id_tmp);
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
                un_id = un_id_tmp;
                un_user_id = un_user_id_tmp;
            }

            return chek_service_user_tmp;
        }

        //<--- სერვისის მომხმარებლის ადმინისტრირება


        //სერვისის ცნობარის მიღება --->

        /// <summary>აქციზური საქონლის კოდები</summary>
        /// <param name="s_text">საქონლის კოდი (თუ "" გადავეცით მაშინ სრული სია, ERROR = The maximum message size quota for incoming messages (65536))</param>
        /// <param name="errorText"></param>
        /// <returns>Dictionary, ან NULL</returns>
        public Dictionary<string, HashSet<string>> get_akciz_codes(string s_text, out string errorText)
        {
            errorText = null;
            Dictionary<string, HashSet<string>> akciz_codes_map = null;
            XmlNode get_akciz_codes_result = null;
            try
            {
                if (protocolType == "HTTP")
                {
                    get_akciz_codes_result = wayBill_soapClient_HTTP.get_akciz_codes(su, sp, s_text);
                }
                else
                {
                    get_akciz_codes_result = wayBill_soapClient_HTTPS.get_akciz_codes(su, sp, s_text);
                }
            }
            catch (Exception ex)
            {
                errorText = ex.Message + "! get_akciz_codes()";
                return akciz_codes_map;
            }

            if (get_akciz_codes_result != null)
            {
                akciz_codes_map = new Dictionary<string, HashSet<string>>();

                XmlNodeList itemNodes = get_akciz_codes_result.SelectNodes("AKCIZ_CODE");
                foreach (XmlNode itemNode in itemNodes)
                {
                    string ID = (itemNode.SelectSingleNode("ID") == null) ? "" : itemNode.SelectSingleNode("ID").InnerText;
                    string TITLE = (itemNode.SelectSingleNode("TITLE") == null) ? "" : itemNode.SelectSingleNode("TITLE").InnerText;
                    string MEASUREMENT = (itemNode.SelectSingleNode("MEASUREMENT") == null) ? "" : itemNode.SelectSingleNode("MEASUREMENT").InnerText;
                    string SAKON_KODI = (itemNode.SelectSingleNode("SAKON_KODI") == null) ? "" : itemNode.SelectSingleNode("SAKON_KODI").InnerText;
                    string AKCIS_GANAKV = (itemNode.SelectSingleNode("AKCIS_GANAKV") == null) ? "" : itemNode.SelectSingleNode("AKCIS_GANAKV").InnerText;
                    akciz_codes_map.Add(ID, new HashSet<string>() { TITLE, MEASUREMENT, SAKON_KODI, AKCIS_GANAKV });
                }
            }

            return akciz_codes_map;
        }

        /// <summary>ზედნადებების ტიპების გამოტანა</summary>
        /// <param name="errorText"></param>
        /// <returns>Dictionary, ან NULL</returns>
        public Dictionary<string, string> get_waybill_types(out string errorText)
        {
            errorText = null;
            Dictionary<string, string> waybill_types_map = null;
            XmlNode get_waybill_types_result = null;
            try
            {
                if (protocolType == "HTTP")
                {
                    get_waybill_types_result = wayBill_soapClient_HTTP.get_waybill_types(su, sp);
                }
                else
                {
                    get_waybill_types_result = wayBill_soapClient_HTTPS.get_waybill_types(su, sp);
                }
            }
            catch (Exception ex)
            {
                errorText = ex.Message + "! get_waybill_types()";
                return waybill_types_map;
            }

            if (get_waybill_types_result != null)
            {
                waybill_types_map = new Dictionary<string, string>();

                XmlNodeList itemNodes = get_waybill_types_result.SelectNodes("WAYBILL_TYPE");
                foreach (XmlNode itemNode in itemNodes)
                {
                    string ID = (itemNode.SelectSingleNode("ID") == null) ? "" : itemNode.SelectSingleNode("ID").InnerText;
                    string NAME = (itemNode.SelectSingleNode("NAME") == null) ? "" : itemNode.SelectSingleNode("NAME").InnerText;
                    waybill_types_map.Add(ID, NAME);
                }
            }

            return waybill_types_map;
        }

        /// <summary>საქონლის ერთეულების გამოტანა</summary>
        /// <param name="errorText"></param>
        /// <returns>Dictionary, ან NULL</returns>
        public Dictionary<string, string> get_waybill_units(out string errorText)
        {
            errorText = null;
            Dictionary<string, string> waybill_units_map = null;
            XmlNode get_waybill_units_result = null;
            try
            {
                if (protocolType == "HTTP")
                {
                    get_waybill_units_result = wayBill_soapClient_HTTP.get_waybill_units(su, sp);
                }
                else
                {
                    get_waybill_units_result = wayBill_soapClient_HTTPS.get_waybill_units(su, sp);
                }
            }
            catch (Exception ex)
            {
                errorText = ex.Message + "! get_waybill_units()";
                return waybill_units_map;
            }

            if (get_waybill_units_result != null)
            {
                waybill_units_map = new Dictionary<string, string>();

                XmlNodeList itemNodes = get_waybill_units_result.SelectNodes("WAYBILL_UNIT");
                foreach (XmlNode itemNode in itemNodes)
                {
                    string ID = (itemNode.SelectSingleNode("ID") == null) ? "" : itemNode.SelectSingleNode("ID").InnerText;
                    string NAME = (itemNode.SelectSingleNode("NAME") == null) ? "" : itemNode.SelectSingleNode("NAME").InnerText;
                    waybill_units_map.Add(ID, NAME);
                }
            }

            return waybill_units_map;
        }

        public string get_waybill_unit_name_by_code(string unit_code)
        {
            try
            {
                string errorText;
                Dictionary<string, string> RSUnits = get_waybill_units(out errorText);
                KeyValuePair<string, string> temp = RSUnits.Where(x => x.Key.Equals(unit_code)).FirstOrDefault(); //Contains
                return temp.Value;
            }
            catch
            {
                return null;
            }
        }

        /// <summary>ტრანსპორტირების ტიპის გამოტანა</summary>
        /// <param name="errorText"></param>
        /// <returns>Dictionary, ან NULL</returns>
        public Dictionary<string, string> get_trans_types(out string errorText)
        {
            errorText = null;
            Dictionary<string, string> trans_types_map = null;
            XmlNode get_trans_types_result = null;
            try
            {
                if (protocolType == "HTTP")
                {
                    get_trans_types_result = wayBill_soapClient_HTTP.get_trans_types(su, sp);
                }
                else
                {
                    get_trans_types_result = wayBill_soapClient_HTTPS.get_trans_types(su, sp);
                }
            }
            catch (Exception ex)
            {
                errorText = ex.Message + "! get_trans_types()";
                return trans_types_map;
            }

            if (get_trans_types_result != null)
            {
                trans_types_map = new Dictionary<string, string>();

                XmlNodeList itemNodes = get_trans_types_result.SelectNodes("TRANSPORT_TYPE");
                foreach (XmlNode itemNode in itemNodes)
                {
                    string ID = (itemNode.SelectSingleNode("ID") == null) ? "" : itemNode.SelectSingleNode("ID").InnerText;
                    string NAME = (itemNode.SelectSingleNode("NAME") == null) ? "" : itemNode.SelectSingleNode("NAME").InnerText;
                    trans_types_map.Add(ID, NAME);
                }
            }

            return trans_types_map;
        }

        //<--- სერვისის ცნობარის მიღება


        //სერვისის ელ. ზედნადების წარმოების მეთოდები --->

        /// <summary>ზედნადების შენახვა</summary>
        /// <param name="array_HEADER">ქუდის რეკვიზიტების მასივი</param>
        /// <param name="array_GOODS">ცხრილური ნაწილის რეკვიზიტების მასივი</param>
        /// <param name="errorText"></param>
        /// <returns>თუ წარმატებით გაიგზავნა აბრუნებს 1, თუ არა -1</returns>
        public int save_waybill(string[] array_HEADER, string[][] array_GOODS, out string errorText)
        {
            errorText = null;
            int save_waybill_result_int = -1;
            XmlNode XML_for_save_waybill = createXML_for_save_waybill(array_HEADER, array_GOODS);
            XmlNode save_waybill_result = null;

            try
            {
                if (protocolType == "HTTP")
                {
                    save_waybill_result = wayBill_soapClient_HTTP.save_waybill(su, sp, XML_for_save_waybill);
                }
                else
                {
                    save_waybill_result = wayBill_soapClient_HTTPS.save_waybill(su, sp, XML_for_save_waybill);
                }
            }
            catch (Exception ex)
            {
                errorText = ex.Message;
                return save_waybill_result_int;
            }

            if (save_waybill_result != null)
            {

                string STATUS = (save_waybill_result.SelectSingleNode("STATUS") == null) ? "" : save_waybill_result.SelectSingleNode("STATUS").InnerText;
                string TYPE = "";
                string get_error_codes_result = null;

                if (STATUS == "0")
                {
                    array_HEADER[0] = (save_waybill_result.SelectSingleNode("ID") == null) ? "" : save_waybill_result.SelectSingleNode("ID").InnerText;
                    array_HEADER[19] = (save_waybill_result.SelectSingleNode("WAYBILL_NUMBER") == null) ? "" : save_waybill_result.SelectSingleNode("WAYBILL_NUMBER").InnerText;
                    array_HEADER[27] = STATUS;
                    save_waybill_result_int = 1;
                }
                else
                {
                    get_error_codes_result = get_error_codes(STATUS, TYPE, out errorText);
                    if (errorText == null)
                    {
                        array_HEADER[27] = get_error_codes_result;
                    }
                    else
                    {
                        return save_waybill_result_int;
                    }
                }

                int i = 0;
                string ERROR;

                XmlNodeList itemNodes = save_waybill_result.SelectNodes("GOODS_LIST/GOODS");
                if (itemNodes.Count != 0)
                {
                    foreach (XmlNode itemNode in itemNodes)
                    {
                        ERROR = (itemNode.SelectSingleNode("ERROR") == null) ? "" : itemNode.SelectSingleNode("ERROR").InnerText;

                        if (ERROR == "0")
                        {
                            array_GOODS[i][0] = (itemNode.SelectSingleNode("ID") == null) ? "" : itemNode.SelectSingleNode("ID").InnerText;
                            array_GOODS[i][12] = ERROR;
                        }
                        else
                        {
                            get_error_codes_result = get_error_codes(STATUS, TYPE, out errorText);
                            if (errorText == null)
                            {
                                array_GOODS[i][12] = get_error_codes_result;
                            }
                        }

                        i++;
                    }
                }
            }

            return save_waybill_result_int;
        }

        /// <summary>ქმნის XML-ს ზედნადების გაგზავნისთვის</summary>
        /// <param name="array_HEADER">ქუდის რეკვიზიტების მასივი</param>
        /// <param name="array_GOODS">ცხრილური ნაწილის რეკვიზიტების მასივი</param>
        /// <returns>XML</returns>
        public XmlNode createXML_for_save_waybill(string[] array_HEADER, string[][] array_GOODS)
        {

            XmlDocument xmlDoc = new XmlDocument();
            XmlNode WAYBILL = xmlDoc.CreateElement("WAYBILL");
            WAYBILL.AppendChild(xmlDoc.CreateElement("SUB_WAYBILLS"));
            XmlNode GOODS_LIST = xmlDoc.CreateElement("GOODS_LIST");
            WAYBILL.AppendChild(GOODS_LIST);
            int array_Length = array_GOODS.Length;

            for (int i = 0; i < array_Length; i++)
            {
                XmlNode GOODS = xmlDoc.CreateElement("GOODS");
                XmlNode ID = xmlDoc.CreateElement("ID");
                ID.InnerText = array_GOODS[i][0];
                GOODS.AppendChild(ID);
                XmlNode W_NAME = xmlDoc.CreateElement("W_NAME");
                W_NAME.InnerText = array_GOODS[i][1];
                GOODS.AppendChild(W_NAME);
                XmlNode UNIT_ID = xmlDoc.CreateElement("UNIT_ID");
                UNIT_ID.InnerText = array_GOODS[i][2];
                GOODS.AppendChild(UNIT_ID);
                XmlNode UNIT_TXT = xmlDoc.CreateElement("UNIT_TXT");
                UNIT_TXT.InnerText = array_GOODS[i][3];
                GOODS.AppendChild(UNIT_TXT);
                XmlNode QUANTITY = xmlDoc.CreateElement("QUANTITY");
                QUANTITY.InnerText = array_GOODS[i][4];
                GOODS.AppendChild(QUANTITY);
                XmlNode PRICE = xmlDoc.CreateElement("PRICE");
                PRICE.InnerText = array_GOODS[i][5];
                GOODS.AppendChild(PRICE);
                XmlNode STATUS = xmlDoc.CreateElement("STATUS");
                STATUS.InnerText = array_GOODS[i][6];
                GOODS.AppendChild(STATUS);
                XmlNode AMOUNT = xmlDoc.CreateElement("AMOUNT");
                AMOUNT.InnerText = array_GOODS[i][7];
                GOODS.AppendChild(AMOUNT);
                XmlNode BAR_CODE = xmlDoc.CreateElement("BAR_CODE");
                BAR_CODE.InnerText = array_GOODS[i][8];
                GOODS.AppendChild(BAR_CODE);
                XmlNode A_ID = xmlDoc.CreateElement("A_ID");
                A_ID.InnerText = array_GOODS[i][9];
                GOODS.AppendChild(A_ID);
                XmlNode VAT_TYPE = xmlDoc.CreateElement("VAT_TYPE");
                VAT_TYPE.InnerText = array_GOODS[i][10];
                GOODS.AppendChild(VAT_TYPE);
                XmlNode QUANTITY_EXT = xmlDoc.CreateElement("QUANTITY_EXT");
                QUANTITY_EXT.InnerText = array_GOODS[i][11];
                GOODS.AppendChild(QUANTITY_EXT);

                GOODS_LIST.AppendChild(GOODS);
            }

            XmlNode ID_HEADER = xmlDoc.CreateElement("ID");
            ID_HEADER.InnerText = array_HEADER[0];
            WAYBILL.AppendChild(ID_HEADER);
            XmlNode TYPE = xmlDoc.CreateElement("TYPE");
            TYPE.InnerText = array_HEADER[1];
            WAYBILL.AppendChild(TYPE);
            XmlNode BUYER_TIN = xmlDoc.CreateElement("BUYER_TIN");
            BUYER_TIN.InnerText = array_HEADER[2];
            WAYBILL.AppendChild(BUYER_TIN);
            XmlNode CHEK_BUYER_TIN = xmlDoc.CreateElement("CHEK_BUYER_TIN");
            CHEK_BUYER_TIN.InnerText = array_HEADER[3];
            WAYBILL.AppendChild(CHEK_BUYER_TIN);
            XmlNode BUYER_NAME = xmlDoc.CreateElement("BUYER_NAME");
            BUYER_NAME.InnerText = array_HEADER[4];
            WAYBILL.AppendChild(BUYER_NAME);
            XmlNode START_ADDRESS = xmlDoc.CreateElement("START_ADDRESS");
            START_ADDRESS.InnerText = array_HEADER[5];
            WAYBILL.AppendChild(START_ADDRESS);
            XmlNode END_ADDRESS = xmlDoc.CreateElement("END_ADDRESS");
            END_ADDRESS.InnerText = array_HEADER[6];
            WAYBILL.AppendChild(END_ADDRESS);
            XmlNode DRIVER_TIN = xmlDoc.CreateElement("DRIVER_TIN");
            DRIVER_TIN.InnerText = array_HEADER[7];
            WAYBILL.AppendChild(DRIVER_TIN);
            XmlNode CHEK_DRIVER_TIN = xmlDoc.CreateElement("CHEK_DRIVER_TIN");
            CHEK_DRIVER_TIN.InnerText = array_HEADER[8];
            WAYBILL.AppendChild(CHEK_DRIVER_TIN);
            XmlNode DRIVER_NAME = xmlDoc.CreateElement("DRIVER_NAME");
            DRIVER_NAME.InnerText = array_HEADER[9];
            WAYBILL.AppendChild(DRIVER_NAME);
            XmlNode TRANSPORT_COAST = xmlDoc.CreateElement("TRANSPORT_COAST");
            TRANSPORT_COAST.InnerText = array_HEADER[10];
            WAYBILL.AppendChild(TRANSPORT_COAST);
            XmlNode RECEPTION_INFO = xmlDoc.CreateElement("RECEPTION_INFO");
            RECEPTION_INFO.InnerText = array_HEADER[11];
            WAYBILL.AppendChild(RECEPTION_INFO);
            XmlNode RECEIVER_INFO = xmlDoc.CreateElement("RECEIVER_INFO");
            RECEIVER_INFO.InnerText = array_HEADER[12];
            WAYBILL.AppendChild(RECEIVER_INFO);
            XmlNode DELIVERY_DATE = xmlDoc.CreateElement("DELIVERY_DATE");
            DELIVERY_DATE.InnerText = array_HEADER[13];
            WAYBILL.AppendChild(DELIVERY_DATE);
            XmlNode STATUS_HEADER = xmlDoc.CreateElement("STATUS");
            STATUS_HEADER.InnerText = array_HEADER[14];
            WAYBILL.AppendChild(STATUS_HEADER);
            XmlNode SELER_UN_ID = xmlDoc.CreateElement("SELER_UN_ID");
            SELER_UN_ID.InnerText = array_HEADER[15];
            WAYBILL.AppendChild(SELER_UN_ID);
            XmlNode PAR_ID = xmlDoc.CreateElement("PAR_ID");
            PAR_ID.InnerText = array_HEADER[16];
            WAYBILL.AppendChild(PAR_ID);
            XmlNode FULL_AMOUNT = xmlDoc.CreateElement("FULL_AMOUNT");
            FULL_AMOUNT.InnerText = array_HEADER[17];
            WAYBILL.AppendChild(FULL_AMOUNT);
            XmlNode CAR_NUMBER = xmlDoc.CreateElement("CAR_NUMBER");
            CAR_NUMBER.InnerText = array_HEADER[18];
            WAYBILL.AppendChild(CAR_NUMBER);
            XmlNode WAYBILL_NUMBER = xmlDoc.CreateElement("WAYBILL_NUMBER");
            WAYBILL_NUMBER.InnerText = array_HEADER[19];
            WAYBILL.AppendChild(WAYBILL_NUMBER);
            XmlNode S_USER_ID = xmlDoc.CreateElement("S_USER_ID");
            S_USER_ID.InnerText = array_HEADER[20];
            WAYBILL.AppendChild(S_USER_ID);
            XmlNode BEGIN_DATE = xmlDoc.CreateElement("BEGIN_DATE");
            BEGIN_DATE.InnerText = array_HEADER[21];
            WAYBILL.AppendChild(BEGIN_DATE);
            XmlNode TRAN_COST_PAYER = xmlDoc.CreateElement("TRAN_COST_PAYER");
            TRAN_COST_PAYER.InnerText = array_HEADER[22];
            WAYBILL.AppendChild(TRAN_COST_PAYER);
            XmlNode TRANS_ID = xmlDoc.CreateElement("TRANS_ID");
            TRANS_ID.InnerText = array_HEADER[23];
            WAYBILL.AppendChild(TRANS_ID);
            XmlNode TRANS_TXT = xmlDoc.CreateElement("TRANS_TXT");
            TRANS_TXT.InnerText = array_HEADER[24];
            WAYBILL.AppendChild(TRANS_TXT);
            XmlNode COMMENT = xmlDoc.CreateElement("COMMENT");
            COMMENT.InnerText = array_HEADER[25];
            WAYBILL.AppendChild(COMMENT);
            XmlNode TRANSPORTER_TIN = xmlDoc.CreateElement("TRANSPORTER_TIN");
            TRANSPORTER_TIN.InnerText = array_HEADER[26];
            WAYBILL.AppendChild(TRANSPORTER_TIN);

            return WAYBILL;
        }

        /// <summary>ზედნადების გამოტანა (by ID)</summary>
        /// <param name="waybill_id">ზედნადების ID</param>
        /// <param name="array_HEADER">ქუდის რეკვიზიტების მასივი</param>
        /// <param name="array_GOODS">ცხრილური ნაწილის რეკვიზიტების მასივი</param>
        /// <param name="arry_SUB_WAYBILLS">ქვეზედნადების მასივი</param>
        /// <param name="errorText"></param>
        /// <returns>თუ წარმატებით გამოიტანა აბრუნებს 1, თუ არა -1</returns>
        public int get_waybill(int waybill_id, out string[] array_HEADER, out string[][] array_GOODS, out string[][] arry_SUB_WAYBILLS, out string errorText)
        {
            errorText = null;
            XmlNode get_waybill_result = null;
            int get_waybill_result_int = -1;
            array_HEADER = null;
            array_GOODS = null;
            arry_SUB_WAYBILLS = null;
            try
            {
                if (protocolType == "HTTP")
                {
                    get_waybill_result = wayBill_soapClient_HTTP.get_waybill(su, sp, waybill_id);
                }
                else
                {
                    get_waybill_result = wayBill_soapClient_HTTPS.get_waybill(su, sp, waybill_id);
                }
            }
            catch (Exception ex)
            {
                errorText = ex.Message + "! get_waybill()";
                return get_waybill_result_int;
            }

            if (get_waybill_result != null)
            {
                array_HEADER = new string[46];
                array_HEADER[0] = (get_waybill_result.SelectSingleNode("ID") == null) ? "" : get_waybill_result.SelectSingleNode("ID").InnerText;
                array_HEADER[1] = (get_waybill_result.SelectSingleNode("TYPE") == null) ? "" : get_waybill_result.SelectSingleNode("TYPE").InnerText;
                array_HEADER[2] = (get_waybill_result.SelectSingleNode("CREATE_DATE") == null) ? "" : get_waybill_result.SelectSingleNode("CREATE_DATE").InnerText;
                array_HEADER[3] = (get_waybill_result.SelectSingleNode("BUYER_TIN") == null) ? "" : get_waybill_result.SelectSingleNode("BUYER_TIN").InnerText;
                array_HEADER[4] = (get_waybill_result.SelectSingleNode("CHEK_BUYER_TIN") == null) ? "" : get_waybill_result.SelectSingleNode("CHEK_BUYER_TIN").InnerText;
                array_HEADER[5] = (get_waybill_result.SelectSingleNode("BUYER_NAME") == null) ? "" : get_waybill_result.SelectSingleNode("BUYER_NAME").InnerText;
                array_HEADER[6] = (get_waybill_result.SelectSingleNode("START_ADDRESS") == null) ? "" : get_waybill_result.SelectSingleNode("START_ADDRESS").InnerText;
                array_HEADER[7] = (get_waybill_result.SelectSingleNode("END_ADDRESS") == null) ? "" : get_waybill_result.SelectSingleNode("END_ADDRESS").InnerText;
                array_HEADER[8] = (get_waybill_result.SelectSingleNode("DRIVER_TIN") == null) ? "" : get_waybill_result.SelectSingleNode("DRIVER_TIN").InnerText;
                array_HEADER[9] = (get_waybill_result.SelectSingleNode("CHEK_DRIVER_TIN") == null) ? "" : get_waybill_result.SelectSingleNode("CHEK_DRIVER_TIN").InnerText;
                array_HEADER[10] = (get_waybill_result.SelectSingleNode("DRIVER_NAME") == null) ? "" : get_waybill_result.SelectSingleNode("DRIVER_NAME").InnerText;
                array_HEADER[11] = (get_waybill_result.SelectSingleNode("TRANSPORT_COAST") == null) ? "" : get_waybill_result.SelectSingleNode("TRANSPORT_COAST").InnerText;
                array_HEADER[12] = (get_waybill_result.SelectSingleNode("RECEPTION_INFO") == null) ? "" : get_waybill_result.SelectSingleNode("RECEPTION_INFO").InnerText;
                array_HEADER[13] = (get_waybill_result.SelectSingleNode("RECEIVER_INFO") == null) ? "" : get_waybill_result.SelectSingleNode("RECEIVER_INFO").InnerText;
                array_HEADER[14] = (get_waybill_result.SelectSingleNode("DELIVERY_DATE") == null) ? "" : get_waybill_result.SelectSingleNode("DELIVERY_DATE").InnerText;
                array_HEADER[15] = (get_waybill_result.SelectSingleNode("STATUS") == null) ? "" : get_waybill_result.SelectSingleNode("STATUS").InnerText;
                array_HEADER[16] = (get_waybill_result.SelectSingleNode("SELER_UN_ID") == null) ? "" : get_waybill_result.SelectSingleNode("SELER_UN_ID").InnerText;
                array_HEADER[17] = (get_waybill_result.SelectSingleNode("ACTIVATE_DATE") == null) ? "" : get_waybill_result.SelectSingleNode("ACTIVATE_DATE").InnerText;
                array_HEADER[18] = (get_waybill_result.SelectSingleNode("PAR_ID") == null) ? "" : get_waybill_result.SelectSingleNode("PAR_ID").InnerText;
                array_HEADER[19] = (get_waybill_result.SelectSingleNode("FULL_AMOUNT") == null) ? "" : get_waybill_result.SelectSingleNode("FULL_AMOUNT").InnerText;
                array_HEADER[20] = (get_waybill_result.SelectSingleNode("FULL_AMOUNT_TXT") == null) ? "" : get_waybill_result.SelectSingleNode("FULL_AMOUNT_TXT").InnerText;
                array_HEADER[21] = (get_waybill_result.SelectSingleNode("CAR_NUMBER") == null) ? "" : get_waybill_result.SelectSingleNode("CAR_NUMBER").InnerText;
                array_HEADER[22] = (get_waybill_result.SelectSingleNode("WAYBILL_NUMBER") == null) ? "" : get_waybill_result.SelectSingleNode("WAYBILL_NUMBER").InnerText;
                array_HEADER[23] = (get_waybill_result.SelectSingleNode("CLOSE_DATE") == null) ? "" : get_waybill_result.SelectSingleNode("CLOSE_DATE").InnerText;
                array_HEADER[24] = (get_waybill_result.SelectSingleNode("S_USER_ID") == null) ? "" : get_waybill_result.SelectSingleNode("S_USER_ID").InnerText;
                array_HEADER[25] = (get_waybill_result.SelectSingleNode("BEGIN_DATE") == null) ? "" : get_waybill_result.SelectSingleNode("BEGIN_DATE").InnerText;
                array_HEADER[26] = (get_waybill_result.SelectSingleNode("TRAN_COST_PAYER") == null) ? "" : get_waybill_result.SelectSingleNode("TRAN_COST_PAYER").InnerText;
                array_HEADER[27] = (get_waybill_result.SelectSingleNode("TRANS_ID") == null) ? "" : get_waybill_result.SelectSingleNode("TRANS_ID").InnerText;
                array_HEADER[28] = (get_waybill_result.SelectSingleNode("TRANS_TXT") == null) ? "" : get_waybill_result.SelectSingleNode("TRANS_TXT").InnerText;
                array_HEADER[29] = (get_waybill_result.SelectSingleNode("COMMENT") == null) ? "" : get_waybill_result.SelectSingleNode("COMMENT").InnerText;
                array_HEADER[30] = (get_waybill_result.SelectSingleNode("IS_CONFIRMED") == null) ? "" : get_waybill_result.SelectSingleNode("IS_CONFIRMED").InnerText;
                array_HEADER[31] = (get_waybill_result.SelectSingleNode("INVOICE_ID") == null) ? "" : get_waybill_result.SelectSingleNode("INVOICE_ID").InnerText;
                array_HEADER[32] = (get_waybill_result.SelectSingleNode("CONFIRMATION_DATE") == null) ? "" : get_waybill_result.SelectSingleNode("CONFIRMATION_DATE").InnerText;
                array_HEADER[33] = (get_waybill_result.SelectSingleNode("SELLER_TIN") == null) ? "" : get_waybill_result.SelectSingleNode("SELLER_TIN").InnerText;
                array_HEADER[34] = (get_waybill_result.SelectSingleNode("SELLER_NAME") == null) ? "" : get_waybill_result.SelectSingleNode("SELLER_NAME").InnerText;
                array_HEADER[35] = (get_waybill_result.SelectSingleNode("WOOD_LABELS") == null) ? "" : get_waybill_result.SelectSingleNode("WOOD_LABELS").InnerText;
                array_HEADER[36] = (get_waybill_result.SelectSingleNode("CATEGORY") == null) ? "" : get_waybill_result.SelectSingleNode("CATEGORY").InnerText;
                array_HEADER[37] = (get_waybill_result.SelectSingleNode("ORIGIN_TYPE") == null) ? "" : get_waybill_result.SelectSingleNode("ORIGIN_TYPE").InnerText;
                array_HEADER[38] = (get_waybill_result.SelectSingleNode("ORIGIN_TEXT") == null) ? "" : get_waybill_result.SelectSingleNode("ORIGIN_TEXT").InnerText;
                array_HEADER[39] = (get_waybill_result.SelectSingleNode("BUYER_S_USER_ID") == null) ? "" : get_waybill_result.SelectSingleNode("BUYER_S_USER_ID").InnerText;
                array_HEADER[40] = (get_waybill_result.SelectSingleNode("TOTAL_QUANTITY") == null) ? "" : get_waybill_result.SelectSingleNode("TOTAL_QUANTITY").InnerText;
                array_HEADER[41] = (get_waybill_result.SelectSingleNode("TRANSPORTER_TIN") == null) ? "" : get_waybill_result.SelectSingleNode("TRANSPORTER_TIN").InnerText;
                array_HEADER[42] = (get_waybill_result.SelectSingleNode("CUST_STATUS") == null) ? "" : get_waybill_result.SelectSingleNode("CUST_STATUS").InnerText;
                array_HEADER[43] = (get_waybill_result.SelectSingleNode("CUST_NAME") == null) ? "" : get_waybill_result.SelectSingleNode("RECEIVER_INFO").InnerText;

                double QUANTITY = 0;
                double AMOUNT = 0;

                XmlNodeList itemNodes = get_waybill_result.SelectNodes("GOODS_LIST/GOODS");
                if (itemNodes.Count != 0)
                {
                    int i = 0;
                    int size = itemNodes.Count;
                    array_GOODS = new string[size][];
                    Dictionary<string, string> activeDimensionsList = CommonFunctions.getActiveDimensionsList(out errorText); // array-ს ზომა რომ სწორად დავსვათ იმისთვის

                    foreach (XmlNode itemNode in itemNodes)
                    {
                        array_GOODS[i] = new string[14+activeDimensionsList.Count];
                        array_GOODS[i][0] = (itemNode.SelectSingleNode("ID") == null) ? "" : itemNode.SelectSingleNode("ID").InnerText;
                        array_GOODS[i][1] = (itemNode.SelectSingleNode("W_NAME") == null) ? "" : itemNode.SelectSingleNode("W_NAME").InnerText;
                        array_GOODS[i][2] = (itemNode.SelectSingleNode("UNIT_ID") == null) ? "" : itemNode.SelectSingleNode("UNIT_ID").InnerText;
                        array_GOODS[i][3] = (itemNode.SelectSingleNode("QUANTITY") == null) ? "" : itemNode.SelectSingleNode("QUANTITY").InnerText;
                        array_GOODS[i][4] = (itemNode.SelectSingleNode("PRICE") == null) ? "" : itemNode.SelectSingleNode("PRICE").InnerText;
                        array_GOODS[i][5] = (itemNode.SelectSingleNode("AMOUNT") == null) ? "" : itemNode.SelectSingleNode("AMOUNT").InnerText;
                        array_GOODS[i][6] = (itemNode.SelectSingleNode("BAR_CODE") == null) ? "" : itemNode.SelectSingleNode("BAR_CODE").InnerText;
                        array_GOODS[i][7] = (itemNode.SelectSingleNode("A_ID") == null) ? "" : itemNode.SelectSingleNode("A_ID").InnerText;
                        array_GOODS[i][8] = (itemNode.SelectSingleNode("VAT_TYPE") == null) ? "" : itemNode.SelectSingleNode("VAT_TYPE").InnerText;
                        array_GOODS[i][9] = (itemNode.SelectSingleNode("QUANTITY_EXT") == null) ? "" : itemNode.SelectSingleNode("QUANTITY_EXT").InnerText;
                        array_GOODS[i][10] = (itemNode.SelectSingleNode("STATUS") == null) ? "" : itemNode.SelectSingleNode("STATUS").InnerText;
                        array_GOODS[i][11] = (itemNode.SelectSingleNode("QUANTITY_F") == null) ? "" : itemNode.SelectSingleNode("QUANTITY_F").InnerText;
                        array_GOODS[i][12] = "";
                        array_GOODS[i][13] = (itemNode.SelectSingleNode("UNIT_TXT") == null) ? "" : itemNode.SelectSingleNode("UNIT_TXT").InnerText;
                        QUANTITY = QUANTITY + (array_GOODS[i][3] == "" ? 0 : Convert.ToDouble(array_GOODS[i][3], CultureInfo.InvariantCulture));
                        AMOUNT = AMOUNT + (array_GOODS[i][5] == "" ? 0 : Convert.ToDouble(array_GOODS[i][5], CultureInfo.InvariantCulture));

                        i++;
                    }
                }

                NumberFormatInfo Nfi = new NumberFormatInfo() { NumberDecimalSeparator = "." };
                array_HEADER[44] = QUANTITY.ToString(Nfi);
                array_HEADER[45] = AMOUNT.ToString(Nfi);

                //<WOOD_DOCS_LIST></WOOD_DOCS_LIST>
                itemNodes = get_waybill_result.SelectNodes("SUB_WAYBILLS/SUB_WAYBILL");
                if (itemNodes.Count != 0)
                {
                    int i = 0;
                    int size = itemNodes.Count;
                    arry_SUB_WAYBILLS = new string[size][];

                    foreach (XmlNode itemNode in itemNodes)
                    {
                        arry_SUB_WAYBILLS[i] = new string[6];
                        arry_SUB_WAYBILLS[i][0] = (itemNode.SelectSingleNode("ID") == null) ? "" : itemNode.SelectSingleNode("ID").InnerText;
                        arry_SUB_WAYBILLS[i][1] = (itemNode.SelectSingleNode("WAYBILL_NUMBER") == null) ? "" : itemNode.SelectSingleNode("WAYBILL_NUMBER").InnerText;
                        arry_SUB_WAYBILLS[i][2] = (itemNode.SelectSingleNode("BUYER_TIN") == null) ? "" : itemNode.SelectSingleNode("BUYER_TIN").InnerText;
                        arry_SUB_WAYBILLS[i][3] = (itemNode.SelectSingleNode("BUYER_NAME") == null) ? "" : itemNode.SelectSingleNode("BUYER_NAME").InnerText;
                        arry_SUB_WAYBILLS[i][4] = (itemNode.SelectSingleNode("FULL_AMOUNT") == null) ? "" : itemNode.SelectSingleNode("FULL_AMOUNT").InnerText;
                        arry_SUB_WAYBILLS[i][5] = (itemNode.SelectSingleNode("STATUS") == null) ? "" : itemNode.SelectSingleNode("STATUS").InnerText;

                        i++;
                    }
                }

                get_waybill_result_int = 1;
            }

            return get_waybill_result_int;
        }

        /// <summary>ზედნადების გამოტანა (by Number)</summary>
        /// <param name="waybill_number">ზედნადების ნომერი</param>
        /// <param name="array_HEADER">ქუდის რეკვიზიტების მასივი</param>
        /// <param name="array_GOODS">ცხრილური ნაწილის რეკვიზიტების მასივი</param>
        /// <param name="arry_SUB_WAYBILLS">ქვეზედნადების მასივი</param>
        /// <param name="errorText"></param>
        /// <returns>თუ წარმატებით გამოიტანა აბრუნებს 1, თუ არა -1</returns>
        public int get_waybill_by_number(string waybill_number, out string[] array_HEADER, out string[][] array_GOODS, out string[][] arry_SUB_WAYBILLS, out string errorText)
        {
            errorText = null;
            XmlNode get_waybill_result = null;
            int get_waybill_result_int = -1;
            array_HEADER = null;
            array_GOODS = null;
            arry_SUB_WAYBILLS = null;
            try
            {
                if (protocolType == "HTTP")
                {
                    get_waybill_result = wayBill_soapClient_HTTP.get_waybill_by_number(su, sp, waybill_number);
                }
                else
                {
                    get_waybill_result = wayBill_soapClient_HTTPS.get_waybill_by_number(su, sp, waybill_number);
                }
            }
            catch (Exception ex)
            {
                errorText = ex.Message + "! get_waybill()";
                return get_waybill_result_int;
            }

            if (get_waybill_result != null)
            {
                array_HEADER = new string[34];
                array_HEADER[0] = (get_waybill_result.SelectSingleNode("ID") == null) ? "" : get_waybill_result.SelectSingleNode("ID").InnerText;
                array_HEADER[1] = (get_waybill_result.SelectSingleNode("TYPE") == null) ? "" : get_waybill_result.SelectSingleNode("TYPE").InnerText;
                array_HEADER[2] = (get_waybill_result.SelectSingleNode("CREATE_DATE") == null) ? "" : get_waybill_result.SelectSingleNode("CREATE_DATE").InnerText;
                array_HEADER[3] = (get_waybill_result.SelectSingleNode("BUYER_TIN") == null) ? "" : get_waybill_result.SelectSingleNode("BUYER_TIN").InnerText;
                array_HEADER[4] = (get_waybill_result.SelectSingleNode("CHEK_BUYER_TIN") == null) ? "" : get_waybill_result.SelectSingleNode("CHEK_BUYER_TIN").InnerText;
                array_HEADER[5] = (get_waybill_result.SelectSingleNode("BUYER_NAME") == null) ? "" : get_waybill_result.SelectSingleNode("BUYER_NAME").InnerText;
                array_HEADER[6] = (get_waybill_result.SelectSingleNode("START_ADDRESS") == null) ? "" : get_waybill_result.SelectSingleNode("START_ADDRESS").InnerText;
                array_HEADER[7] = (get_waybill_result.SelectSingleNode("END_ADDRESS") == null) ? "" : get_waybill_result.SelectSingleNode("END_ADDRESS").InnerText;
                array_HEADER[8] = (get_waybill_result.SelectSingleNode("DRIVER_TIN") == null) ? "" : get_waybill_result.SelectSingleNode("DRIVER_TIN").InnerText;
                array_HEADER[9] = (get_waybill_result.SelectSingleNode("CHEK_DRIVER_TIN") == null) ? "" : get_waybill_result.SelectSingleNode("CHEK_DRIVER_TIN").InnerText;
                array_HEADER[10] = (get_waybill_result.SelectSingleNode("DRIVER_NAME") == null) ? "" : get_waybill_result.SelectSingleNode("DRIVER_NAME").InnerText;
                array_HEADER[11] = (get_waybill_result.SelectSingleNode("TRANSPORT_COAST") == null) ? "" : get_waybill_result.SelectSingleNode("TRANSPORT_COAST").InnerText;
                array_HEADER[12] = (get_waybill_result.SelectSingleNode("RECEPTION_INFO") == null) ? "" : get_waybill_result.SelectSingleNode("RECEPTION_INFO").InnerText;
                array_HEADER[13] = (get_waybill_result.SelectSingleNode("RECEIVER_INFO") == null) ? "" : get_waybill_result.SelectSingleNode("RECEIVER_INFO").InnerText;
                array_HEADER[14] = (get_waybill_result.SelectSingleNode("DELIVERY_DATE") == null) ? "" : get_waybill_result.SelectSingleNode("DELIVERY_DATE").InnerText;
                array_HEADER[15] = (get_waybill_result.SelectSingleNode("STATUS") == null) ? "" : get_waybill_result.SelectSingleNode("STATUS").InnerText;
                array_HEADER[16] = (get_waybill_result.SelectSingleNode("SELER_UN_ID") == null) ? "" : get_waybill_result.SelectSingleNode("SELER_UN_ID").InnerText;
                array_HEADER[17] = (get_waybill_result.SelectSingleNode("ACTIVATE_DATE") == null) ? "" : get_waybill_result.SelectSingleNode("ACTIVATE_DATE").InnerText;
                array_HEADER[18] = (get_waybill_result.SelectSingleNode("PAR_ID") == null) ? "" : get_waybill_result.SelectSingleNode("PAR_ID").InnerText;
                array_HEADER[19] = (get_waybill_result.SelectSingleNode("FULL_AMOUNT") == null) ? "" : get_waybill_result.SelectSingleNode("FULL_AMOUNT").InnerText;
                array_HEADER[20] = (get_waybill_result.SelectSingleNode("FULL_AMOUNT_TXT") == null) ? "" : get_waybill_result.SelectSingleNode("FULL_AMOUNT_TXT").InnerText;
                array_HEADER[21] = (get_waybill_result.SelectSingleNode("CAR_NUMBER") == null) ? "" : get_waybill_result.SelectSingleNode("CAR_NUMBER").InnerText;
                array_HEADER[22] = (get_waybill_result.SelectSingleNode("WAYBILL_NUMBER") == null) ? "" : get_waybill_result.SelectSingleNode("WAYBILL_NUMBER").InnerText;
                array_HEADER[23] = (get_waybill_result.SelectSingleNode("CLOSE_DATE") == null) ? "" : get_waybill_result.SelectSingleNode("CLOSE_DATE").InnerText;
                array_HEADER[24] = (get_waybill_result.SelectSingleNode("S_USER_ID") == null) ? "" : get_waybill_result.SelectSingleNode("S_USER_ID").InnerText;
                array_HEADER[25] = (get_waybill_result.SelectSingleNode("BEGIN_DATE") == null) ? "" : get_waybill_result.SelectSingleNode("BEGIN_DATE").InnerText;
                array_HEADER[26] = (get_waybill_result.SelectSingleNode("TRAN_COST_PAYER") == null) ? "" : get_waybill_result.SelectSingleNode("TRAN_COST_PAYER").InnerText;
                array_HEADER[27] = (get_waybill_result.SelectSingleNode("TRANS_ID") == null) ? "" : get_waybill_result.SelectSingleNode("TRANS_ID").InnerText;
                array_HEADER[28] = (get_waybill_result.SelectSingleNode("TRANS_TXT") == null) ? "" : get_waybill_result.SelectSingleNode("TRANS_TXT").InnerText;
                array_HEADER[29] = (get_waybill_result.SelectSingleNode("COMMENT") == null) ? "" : get_waybill_result.SelectSingleNode("COMMENT").InnerText;
                array_HEADER[30] = (get_waybill_result.SelectSingleNode("IS_CONFIRMED") == null) ? "" : get_waybill_result.SelectSingleNode("IS_CONFIRMED").InnerText;
                array_HEADER[31] = (get_waybill_result.SelectSingleNode("CONFIRMATION_DATE") == null) ? "" : get_waybill_result.SelectSingleNode("CONFIRMATION_DATE").InnerText;
                array_HEADER[32] = (get_waybill_result.SelectSingleNode("SELLER_TIN") == null) ? "" : get_waybill_result.SelectSingleNode("SELLER_TIN").InnerText;
                array_HEADER[33] = (get_waybill_result.SelectSingleNode("SELLER_NAME") == null) ? "" : get_waybill_result.SelectSingleNode("SELLER_NAME").InnerText;

                XmlNodeList itemNodes = get_waybill_result.SelectNodes("GOODS_LIST/GOODS");
                if (itemNodes.Count != 0)
                {
                    int i = 0;
                    int size = itemNodes.Count;
                    array_GOODS = new string[size][];

                    foreach (XmlNode itemNode in itemNodes)
                    {
                        array_GOODS[i] = new string[11];
                        array_GOODS[i][0] = (itemNode.SelectSingleNode("ID") == null) ? "" : itemNode.SelectSingleNode("ID").InnerText;
                        array_GOODS[i][1] = (itemNode.SelectSingleNode("W_NAME") == null) ? "" : itemNode.SelectSingleNode("W_NAME").InnerText;
                        array_GOODS[i][2] = (itemNode.SelectSingleNode("UNIT_ID") == null) ? "" : itemNode.SelectSingleNode("UNIT_ID").InnerText;
                        array_GOODS[i][3] = (itemNode.SelectSingleNode("QUANTITY") == null) ? "" : itemNode.SelectSingleNode("QUANTITY").InnerText;
                        array_GOODS[i][4] = (itemNode.SelectSingleNode("PRICE") == null) ? "" : itemNode.SelectSingleNode("PRICE").InnerText;
                        array_GOODS[i][5] = (itemNode.SelectSingleNode("AMOUNT") == null) ? "" : itemNode.SelectSingleNode("AMOUNT").InnerText;
                        array_GOODS[i][6] = (itemNode.SelectSingleNode("BAR_CODE") == null) ? "" : itemNode.SelectSingleNode("BAR_CODE").InnerText;
                        array_GOODS[i][7] = (itemNode.SelectSingleNode("A_ID") == null) ? "" : itemNode.SelectSingleNode("A_ID").InnerText;
                        array_GOODS[i][8] = (itemNode.SelectSingleNode("VAT_TYPE") == null) ? "" : itemNode.SelectSingleNode("VAT_TYPE").InnerText;
                        array_GOODS[i][9] = (itemNode.SelectSingleNode("QUANTITY_EXT") == null) ? "" : itemNode.SelectSingleNode("QUANTITY_EXT").InnerText;
                        array_GOODS[i][10] = (itemNode.SelectSingleNode("STATUS") == null) ? "" : itemNode.SelectSingleNode("STATUS").InnerText;

                        i++;
                    }
                }

                itemNodes = get_waybill_result.SelectNodes("SUB_WAYBILLS/SUB_WAYBILL");
                if (itemNodes.Count != 0)
                {
                    int i = 0;
                    int size = itemNodes.Count;
                    arry_SUB_WAYBILLS = new string[size][];

                    foreach (XmlNode itemNode in itemNodes)
                    {
                        arry_SUB_WAYBILLS[i] = new string[6];
                        arry_SUB_WAYBILLS[i][0] = (itemNode.SelectSingleNode("ID") == null) ? "" : itemNode.SelectSingleNode("ID").InnerText;
                        arry_SUB_WAYBILLS[i][1] = (itemNode.SelectSingleNode("WAYBILL_NUMBER") == null) ? "" : itemNode.SelectSingleNode("WAYBILL_NUMBER").InnerText;
                        arry_SUB_WAYBILLS[i][2] = (itemNode.SelectSingleNode("BUYER_TIN") == null) ? "" : itemNode.SelectSingleNode("BUYER_TIN").InnerText;
                        arry_SUB_WAYBILLS[i][3] = (itemNode.SelectSingleNode("BUYER_NAME") == null) ? "" : itemNode.SelectSingleNode("BUYER_NAME").InnerText;
                        arry_SUB_WAYBILLS[i][4] = (itemNode.SelectSingleNode("FULL_AMOUNT") == null) ? "" : itemNode.SelectSingleNode("FULL_AMOUNT").InnerText;
                        arry_SUB_WAYBILLS[i][5] = (itemNode.SelectSingleNode("STATUS") == null) ? "" : itemNode.SelectSingleNode("STATUS").InnerText;

                        i++;
                    }
                }

                get_waybill_result_int = 1;
            }

            return get_waybill_result_int;
        }

        /// <summary>გამყიდველის მიერ გამოწერილი ზედნადებების სია (გამყიდველის მხარე)</summary>
        /// <param name="itypes"></param>
        /// <param name="buyer_tin"></param>
        /// <param name="statuses"></param>
        /// <param name="car_number"></param>
        /// <param name="begin_date_s"></param>
        /// <param name="begin_date_e"></param>
        /// <param name="create_date_s"></param>
        /// <param name="create_date_e"></param>
        /// <param name="driver_tin"></param>
        /// <param name="delivery_date_s"></param>
        /// <param name="delivery_date_e"></param>
        /// <param name="full_amount"></param>
        /// <param name="waybill_number"></param>
        /// <param name="close_date_s"></param>
        /// <param name="close_date_e"></param>
        /// <param name="s_user_ids"></param>
        /// <param name="comment"></param>
        /// <param name="errorText"></param>
        /// <returns></returns>
        public Dictionary<string, Dictionary<string, string>> get_waybills(string itypes, string buyer_tin, string statuses, string car_number, DateTime begin_date_s, DateTime begin_date_e, DateTime create_date_s, DateTime create_date_e, string driver_tin, DateTime delivery_date_s, DateTime delivery_date_e, decimal full_amount, string waybill_number, DateTime close_date_s, DateTime close_date_e, string s_user_id, string comment, out string errorText)
        {
            errorText = null;
            Dictionary<string, Dictionary<string, string>> waybills_map = null;
            XmlNode get_waybills_result = null;

            try
            {
                if (protocolType == "HTTP")
                {
                    get_waybills_result = wayBill_soapClient_HTTP.get_waybills(su, sp, itypes, buyer_tin, statuses, car_number, begin_date_s, begin_date_e, create_date_s, create_date_e, driver_tin, delivery_date_s, delivery_date_e, null, waybill_number, close_date_s, close_date_e, s_user_id, comment);
                }
                else
                {
                    get_waybills_result = wayBill_soapClient_HTTPS.get_waybills(su, sp, itypes, buyer_tin, statuses, car_number, begin_date_s, begin_date_e, create_date_s, create_date_e, driver_tin, delivery_date_s, delivery_date_e, null, waybill_number, close_date_s, close_date_e, s_user_id, comment);
                }
            }

            catch (Exception ex)
            {
                errorText = ex.Message + "! get_waybills()";
            }
            if (get_waybills_result != null)
            {
                waybills_map = new Dictionary<string, Dictionary<string, string>>();
                Dictionary<string, string> waybill_map = null;

                XmlNodeList itemNodes = get_waybills_result.SelectNodes("WAYBILL");
                foreach (XmlNode itemNode in itemNodes)
                {
                    waybill_map = new Dictionary<string, string>();

                    string ID = (itemNode.SelectSingleNode("ID") == null) ? "" : itemNode.SelectSingleNode("ID").InnerText;
                    string TYPE = (itemNode.SelectSingleNode("TYPE") == null) ? "" : itemNode.SelectSingleNode("TYPE").InnerText;
                    string CREATE_DATE = (itemNode.SelectSingleNode("CREATE_DATE") == null) ? "" : itemNode.SelectSingleNode("CREATE_DATE").InnerText;
                    string BUYER_TIN = (itemNode.SelectSingleNode("BUYER_TIN") == null) ? "" : itemNode.SelectSingleNode("BUYER_TIN").InnerText;
                    string BUYER_NAME = (itemNode.SelectSingleNode("BUYER_NAME") == null) ? "" : itemNode.SelectSingleNode("BUYER_NAME").InnerText;
                    string START_ADDRESS = (itemNode.SelectSingleNode("START_ADDRESS") == null) ? "" : itemNode.SelectSingleNode("START_ADDRESS").InnerText;
                    string END_ADDRESS = (itemNode.SelectSingleNode("END_ADDRESS") == null) ? "" : itemNode.SelectSingleNode("END_ADDRESS").InnerText;
                    string DRIVER_TIN = (itemNode.SelectSingleNode("DRIVER_TIN") == null) ? "" : itemNode.SelectSingleNode("DRIVER_TIN").InnerText;
                    string DRIVER_NAME = (itemNode.SelectSingleNode("DRIVER_NAME") == null) ? "" : itemNode.SelectSingleNode("DRIVER_NAME").InnerText;
                    string TRANSPORT_COAST = (itemNode.SelectSingleNode("TRANSPORT_COAST") == null) ? "" : itemNode.SelectSingleNode("TRANSPORT_COAST").InnerText;
                    string RECEPTION_INFO = (itemNode.SelectSingleNode("RECEPTION_INFO") == null) ? "" : itemNode.SelectSingleNode("RECEPTION_INFO").InnerText;
                    string RECEIVER_INFO = (itemNode.SelectSingleNode("RECEIVER_INFO") == null) ? "" : itemNode.SelectSingleNode("RECEIVER_INFO").InnerText;
                    string DELIVERY_DATE = (itemNode.SelectSingleNode("DELIVERY_DATE") == null) ? "" : itemNode.SelectSingleNode("DELIVERY_DATE").InnerText;
                    string STATUS = (itemNode.SelectSingleNode("STATUS") == null) ? "" : itemNode.SelectSingleNode("STATUS").InnerText;
                    string ACTIVATE_DATE = (itemNode.SelectSingleNode("ACTIVATE_DATE") == null) ? "" : itemNode.SelectSingleNode("ACTIVATE_DATE").InnerText;
                    string PAR_ID = (itemNode.SelectSingleNode("PAR_ID") == null) ? "" : itemNode.SelectSingleNode("PAR_ID").InnerText;
                    string FULL_AMOUNT = (itemNode.SelectSingleNode("FULL_AMOUNT") == null) ? "" : itemNode.SelectSingleNode("FULL_AMOUNT").InnerText;
                    string CAR_NUMBER = (itemNode.SelectSingleNode("CAR_NUMBER") == null) ? "" : itemNode.SelectSingleNode("CAR_NUMBER").InnerText;
                    string WAYBILL_NUMBER = (itemNode.SelectSingleNode("WAYBILL_NUMBER") == null) ? "" : itemNode.SelectSingleNode("WAYBILL_NUMBER").InnerText;
                    string CLOSE_DATE = (itemNode.SelectSingleNode("CLOSE_DATE") == null) ? "" : itemNode.SelectSingleNode("CLOSE_DATE").InnerText;
                    string S_USER_ID = (itemNode.SelectSingleNode("S_USER_ID") == null) ? "" : itemNode.SelectSingleNode("S_USER_ID").InnerText;
                    string BEGIN_DATE = (itemNode.SelectSingleNode("BEGIN_DATE") == null) ? "" : itemNode.SelectSingleNode("BEGIN_DATE").InnerText;
                    string WAYBILL_COMMENT = (itemNode.SelectSingleNode("WAYBILL_COMMENT") == null) ? "" : itemNode.SelectSingleNode("WAYBILL_COMMENT").InnerText;
                    string IS_CONFIRMED = (itemNode.SelectSingleNode("IS_CONFIRMED") == null) ? "" : itemNode.SelectSingleNode("IS_CONFIRMED").InnerText;
                    string INVOICE_ID = (itemNode.SelectSingleNode("INVOICE_ID") == null) ? "" : itemNode.SelectSingleNode("INVOICE_ID").InnerText;
                    string IS_CORRECTED = (itemNode.SelectSingleNode("IS_CORRECTED") == null) ? "" : itemNode.SelectSingleNode("IS_CORRECTED").InnerText;
                    string BUYER_ST = (itemNode.SelectSingleNode("BUYER_ST") == null) ? "" : itemNode.SelectSingleNode("BUYER_ST").InnerText;

                    waybill_map.Add("ID", ID);
                    waybill_map.Add("TYPE", TYPE);
                    waybill_map.Add("CREATE_DATE", CREATE_DATE);
                    waybill_map.Add("BUYER_TIN", BUYER_TIN);
                    waybill_map.Add("BUYER_NAME", BUYER_NAME);
                    waybill_map.Add("START_ADDRESS", START_ADDRESS);
                    waybill_map.Add("END_ADDRESS", END_ADDRESS);
                    waybill_map.Add("DRIVER_TIN", DRIVER_TIN);
                    waybill_map.Add("DRIVER_NAME", DRIVER_NAME);
                    waybill_map.Add("TRANSPORT_COAST", TRANSPORT_COAST);
                    waybill_map.Add("RECEPTION_INFO", RECEPTION_INFO);
                    waybill_map.Add("RECEIVER_INFO", RECEIVER_INFO);
                    waybill_map.Add("DELIVERY_DATE", DELIVERY_DATE);
                    waybill_map.Add("STATUS", STATUS);
                    waybill_map.Add("ACTIVATE_DATE", ACTIVATE_DATE);
                    waybill_map.Add("PAR_ID", PAR_ID);
                    waybill_map.Add("FULL_AMOUNT", FULL_AMOUNT);
                    waybill_map.Add("CAR_NUMBER", CAR_NUMBER);
                    waybill_map.Add("WAYBILL_NUMBER", WAYBILL_NUMBER);
                    waybill_map.Add("CLOSE_DATE", CLOSE_DATE);
                    waybill_map.Add("S_USER_ID", S_USER_ID);
                    waybill_map.Add("BEGIN_DATE", BEGIN_DATE);
                    waybill_map.Add("WAYBILL_COMMENT", WAYBILL_COMMENT);
                    waybill_map.Add("IS_CONFIRMED", IS_CONFIRMED);
                    waybill_map.Add("INVOICE_ID", INVOICE_ID);
                    waybill_map.Add("IS_CORRECTED", IS_CORRECTED);
                    waybill_map.Add("BUYER_ST", BUYER_ST);

                    waybills_map.Add(ID, waybill_map);
                }
            }
            return waybills_map;
        }

        /// <summary>მყიდველის მიერ მიღებული ზედნადებების სია (მყიდველის მხარე)</summary>
        /// <param name="itypes"></param>
        /// <param name="seller_tin"></param>
        /// <param name="statuses"></param>
        /// <param name="car_number"></param>
        /// <param name="begin_date_s"></param>
        /// <param name="begin_date_e"></param>
        /// <param name="create_date_s"></param>
        /// <param name="create_date_e"></param>
        /// <param name="driver_tin"></param>
        /// <param name="delivery_date_s"></param>
        /// <param name="delivery_date_e"></param>
        /// <param name="full_amount"></param>
        /// <param name="waybill_number"></param>
        /// <param name="close_date_s"></param>
        /// <param name="close_date_e"></param>
        /// <param name="s_user_ids"></param>
        /// <param name="comment"></param>
        /// <param name="errorText"></param>
        /// <returns></returns>
        public Dictionary<string, Dictionary<string, string>> get_buyer_waybills(string itypes, string seller_tin, string statuses, string car_number, DateTime begin_date_s, DateTime begin_date_e, DateTime create_date_s, DateTime create_date_e, string driver_tin, DateTime delivery_date_s, DateTime delivery_date_e, decimal full_amount, string waybill_number, DateTime close_date_s, DateTime close_date_e, string s_user_id, string comment, string StartAddress, string EndAddress, out string errorText)
        {
            errorText = null;
            Dictionary<string, Dictionary<string, string>> waybills_map = null;
            XmlNode get_waybills_result = null;

            try
            {
                if (protocolType == "HTTP")
                {
                    get_waybills_result = wayBill_soapClient_HTTP.get_buyer_waybills(su, sp, itypes, seller_tin, statuses, car_number, begin_date_s, begin_date_e, null, null, driver_tin, null, null, null, waybill_number, null, null, s_user_id, comment);
                }
                else
                {
                    get_waybills_result = wayBill_soapClient_HTTPS.get_buyer_waybills(su, sp, itypes, seller_tin, statuses, car_number, begin_date_s, begin_date_e, null, null, driver_tin, null, null, null, waybill_number, null, null, s_user_id, comment);
                }
            }

            catch (Exception ex)
            {
                errorText = ex.Message + "! get_buyer_waybills()";
            }
            if (get_waybills_result != null)
            {
                waybills_map = new Dictionary<string, Dictionary<string, string>>();
                Dictionary<string, string> waybill_map = null;

                XmlNodeList itemNodes = get_waybills_result.SelectNodes("WAYBILL");
                foreach (XmlNode itemNode in itemNodes)
                {
                    string START_ADDRESS = (itemNode.SelectSingleNode("START_ADDRESS") == null) ? "" : itemNode.SelectSingleNode("START_ADDRESS").InnerText;
                    string END_ADDRESS = (itemNode.SelectSingleNode("END_ADDRESS") == null) ? "" : itemNode.SelectSingleNode("END_ADDRESS").InnerText;

                    if (StartAddress == "blank" && START_ADDRESS != "") continue;                    
                    if (EndAddress == "blank" && END_ADDRESS != "") continue;

                    if (StartAddress != "" && StartAddress != "blank")
                    {
                        if (StartAddress.StartsWith("*") && StartAddress.EndsWith("*"))
                        {                            
                            if (!START_ADDRESS.Contains(StartAddress.Replace("*", "")))
                                continue;
                        }
                        else if (StartAddress.StartsWith("*"))
                        {                            
                            if (!START_ADDRESS.EndsWith(StartAddress.Replace("*", "")))
                                continue;
                        }
                        else if (StartAddress.EndsWith("*"))
                        {                            
                            if (!START_ADDRESS.StartsWith(StartAddress.Replace("*", "")))
                                continue;
                        }
                        else
                        {                            
                            if (StartAddress.Replace("*", "") != START_ADDRESS)
                                continue;
                        }
                    }
                    if (EndAddress != "" && EndAddress != "blank")
                    {
                        if (EndAddress.StartsWith("*") && EndAddress.EndsWith("*"))
                        {
                            if (!END_ADDRESS.Contains(EndAddress.Replace("*", "")))
                                continue;
                        }
                        else if (EndAddress.StartsWith("*"))
                        {
                            if (!END_ADDRESS.EndsWith(EndAddress.Replace("*", "")))
                                continue;
                        }
                        else if (EndAddress.EndsWith("*"))
                        {
                            if (!END_ADDRESS.StartsWith(EndAddress.Replace("*", "")))
                                continue;
                        }
                        else
                        {
                            if (EndAddress.Replace("*", "") != END_ADDRESS)
                                continue;
                        }
                    }

                           
                   

                    waybill_map = new Dictionary<string, string>();

                    string ID = (itemNode.SelectSingleNode("ID") == null) ? "" : itemNode.SelectSingleNode("ID").InnerText;
                    string TYPE = (itemNode.SelectSingleNode("TYPE") == null) ? "" : itemNode.SelectSingleNode("TYPE").InnerText;
                    string BUYER_TIN = (itemNode.SelectSingleNode("BUYER_TIN") == null) ? "" : itemNode.SelectSingleNode("BUYER_TIN").InnerText;
                    string CREATE_DATE = (itemNode.SelectSingleNode("CREATE_DATE") == null) ? "" : itemNode.SelectSingleNode("CREATE_DATE").InnerText;
                    string BUYER_NAME = (itemNode.SelectSingleNode("BUYER_NAME") == null) ? "" : itemNode.SelectSingleNode("BUYER_NAME").InnerText;
                    string SELLER_NAME = (itemNode.SelectSingleNode("SELLER_NAME") == null) ? "" : itemNode.SelectSingleNode("SELLER_NAME").InnerText;
                    string SELLER_TIN = (itemNode.SelectSingleNode("SELLER_TIN") == null) ? "" : itemNode.SelectSingleNode("SELLER_TIN").InnerText;
                    //string START_ADDRESS = (itemNode.SelectSingleNode("START_ADDRESS") == null) ? "" : itemNode.SelectSingleNode("START_ADDRESS").InnerText;
                    //string END_ADDRESS = (itemNode.SelectSingleNode("END_ADDRESS") == null) ? "" : itemNode.SelectSingleNode("END_ADDRESS").InnerText;
                    string DRIVER_TIN = (itemNode.SelectSingleNode("DRIVER_TIN") == null) ? "" : itemNode.SelectSingleNode("DRIVER_TIN").InnerText;
                    string DRIVER_NAME = (itemNode.SelectSingleNode("DRIVER_NAME") == null) ? "" : itemNode.SelectSingleNode("DRIVER_NAME").InnerText;
                    string TRANSPORT_COAST = (itemNode.SelectSingleNode("TRANSPORT_COAST") == null) ? "" : itemNode.SelectSingleNode("TRANSPORT_COAST").InnerText;
                    string RECEPTION_INFO = (itemNode.SelectSingleNode("RECEPTION_INFO") == null) ? "" : itemNode.SelectSingleNode("RECEPTION_INFO").InnerText;
                    string RECEIVER_INFO = (itemNode.SelectSingleNode("RECEIVER_INFO") == null) ? "" : itemNode.SelectSingleNode("RECEIVER_INFO").InnerText;
                    string DELIVERY_DATE = (itemNode.SelectSingleNode("DELIVERY_DATE") == null) ? "" : itemNode.SelectSingleNode("DELIVERY_DATE").InnerText;
                    string STATUS = (itemNode.SelectSingleNode("STATUS") == null) ? "" : itemNode.SelectSingleNode("STATUS").InnerText;
                    string ACTIVATE_DATE = (itemNode.SelectSingleNode("ACTIVATE_DATE") == null) ? "" : itemNode.SelectSingleNode("ACTIVATE_DATE").InnerText;
                    string PAR_ID = (itemNode.SelectSingleNode("PAR_ID") == null) ? "" : itemNode.SelectSingleNode("PAR_ID").InnerText;
                    string FULL_AMOUNT = (itemNode.SelectSingleNode("FULL_AMOUNT") == null) ? "" : itemNode.SelectSingleNode("FULL_AMOUNT").InnerText;
                    string CAR_NUMBER = (itemNode.SelectSingleNode("CAR_NUMBER") == null) ? "" : itemNode.SelectSingleNode("CAR_NUMBER").InnerText;
                    string WAYBILL_NUMBER = (itemNode.SelectSingleNode("WAYBILL_NUMBER") == null) ? "" : itemNode.SelectSingleNode("WAYBILL_NUMBER").InnerText;
                    string CLOSE_DATE = (itemNode.SelectSingleNode("CLOSE_DATE") == null) ? "" : itemNode.SelectSingleNode("CLOSE_DATE").InnerText;
                    string S_USER_ID = (itemNode.SelectSingleNode("S_USER_ID") == null) ? "" : itemNode.SelectSingleNode("S_USER_ID").InnerText;
                    string BEGIN_DATE = (itemNode.SelectSingleNode("BEGIN_DATE") == null) ? "" : itemNode.SelectSingleNode("BEGIN_DATE").InnerText;
                    string WAYBILL_COMMENT = (itemNode.SelectSingleNode("WAYBILL_COMMENT") == null) ? "" : itemNode.SelectSingleNode("WAYBILL_COMMENT").InnerText;
                    string IS_CONFIRMED = (itemNode.SelectSingleNode("IS_CONFIRMED") == null) ? "" : itemNode.SelectSingleNode("IS_CONFIRMED").InnerText;
                    string INVOICE_ID = (itemNode.SelectSingleNode("INVOICE_ID") == null) ? "" : itemNode.SelectSingleNode("INVOICE_ID").InnerText;
                    string IS_CORRECTED = (itemNode.SelectSingleNode("IS_CORRECTED") == null) ? "" : itemNode.SelectSingleNode("IS_CORRECTED").InnerText;
                    string SELLER_ST = (itemNode.SelectSingleNode("SELLER_ST") == null) ? "" : itemNode.SelectSingleNode("SELLER_ST").InnerText;
                    string UOM_CODE = (itemNode.SelectSingleNode("UOM_CODE") == null) ? "" : itemNode.SelectSingleNode("UOM_CODE").InnerText;

                    waybill_map.Add("ID", ID);
                    waybill_map.Add("TYPE", TYPE);
                    waybill_map.Add("CREATE_DATE", CREATE_DATE);
                    waybill_map.Add("BUYER_TIN", BUYER_TIN);
                    waybill_map.Add("BUYER_NAME", BUYER_NAME);
                    waybill_map.Add("SELLER_NAME", SELLER_NAME);
                    waybill_map.Add("SELLER_TIN", SELLER_TIN);
                    waybill_map.Add("START_ADDRESS", START_ADDRESS);
                    waybill_map.Add("END_ADDRESS", END_ADDRESS);
                    waybill_map.Add("DRIVER_TIN", DRIVER_TIN);
                    waybill_map.Add("DRIVER_NAME", DRIVER_NAME);
                    waybill_map.Add("TRANSPORT_COAST", TRANSPORT_COAST);
                    waybill_map.Add("RECEPTION_INFO", RECEPTION_INFO);
                    waybill_map.Add("RECEIVER_INFO", RECEIVER_INFO);
                    waybill_map.Add("DELIVERY_DATE", DELIVERY_DATE);
                    waybill_map.Add("STATUS", STATUS);
                    waybill_map.Add("ACTIVATE_DATE", ACTIVATE_DATE);
                    waybill_map.Add("PAR_ID", PAR_ID);
                    waybill_map.Add("FULL_AMOUNT", FULL_AMOUNT);
                    waybill_map.Add("CAR_NUMBER", CAR_NUMBER);
                    waybill_map.Add("WAYBILL_NUMBER", WAYBILL_NUMBER);
                    waybill_map.Add("CLOSE_DATE", CLOSE_DATE);
                    waybill_map.Add("S_USER_ID", S_USER_ID);
                    waybill_map.Add("BEGIN_DATE", BEGIN_DATE);
                    waybill_map.Add("WAYBILL_COMMENT", WAYBILL_COMMENT);
                    waybill_map.Add("IS_CONFIRMED", IS_CONFIRMED);
                    waybill_map.Add("INVOICE_ID", INVOICE_ID);
                    waybill_map.Add("IS_CORRECTED", IS_CORRECTED);
                    waybill_map.Add("SELLER_ST", SELLER_ST);
                    waybill_map.Add("UOM_CODE", UOM_CODE);

                    waybills_map.Add(ID, waybill_map);
                }
            }
            return waybills_map;
        }

        /// <summary>ზედნადების აქტივაცია</summary>
        /// <param name="waybill_id">ზედნადების ID</param>
        /// <param name="errorText"></param>
        /// <returns>ზედნადების ნომერი, -1 თუ ვერ გაააქტიურა (სტატუსის გამო), -101 არასწორი ID, -100 არასწორი სერვის მომხმარებლის სახელი ან პაროლი, ან NULL</returns>
        public string send_waybill(int waybill_id, out string errorText)
        {
            errorText = null;
            string waybill_number = null;
            try
            {
                if (protocolType == "HTTP")
                {
                    waybill_number = wayBill_soapClient_HTTP.send_waybill(su, sp, waybill_id);
                }
                else
                {
                    waybill_number = wayBill_soapClient_HTTPS.send_waybill(su, sp, waybill_id);
                }
            }
            catch (Exception ex)
            {
                errorText = ex.Message + "! send_waybill()";
                return waybill_number;
            }

            return waybill_number;
        }

        /// <summary>ზედნადების აქტივაცია (begin_date)</summary>
        /// <param name="begin_date">ტრანსპორტირების დაწყების თარიღი</param>
        /// <param name="waybill_id">ზედნადების ID</param>
        /// <param name="errorText"></param>
        /// <returns>ზედნადების ნომერი,"" ძველი თარიღი(begin_date), -1 თუ ვერ გაააქტიურა (სტატუსის გამო), -101 არასწორი ID, -100 არასწორი სერვის მომხმარებლის სახელი ან პაროლი, ან NULL</returns>
        public string send_waybill_vd(DateTime begin_date, int waybill_id, out string errorText)
        {
            errorText = null;
            string waybill_number = null;
            try
            {
                if (protocolType == "HTTP")
                {
                    waybill_number = wayBill_soapClient_HTTP.send_waybil_vd(su, sp, begin_date, waybill_id);
                }
                else
                {
                    waybill_number = wayBill_soapClient_HTTPS.send_waybil_vd(su, sp, begin_date, waybill_id);
                }
            }
            catch (Exception ex)
            {
                errorText = ex.Message + "! send_waybill_vd()";
                return waybill_number;
            }

            return waybill_number;
        }

        /// <summary>ზედნადების აქტივაცია (transporter)</summary>
        /// <param name="waybill_id">ზედნადების ID</param>
        /// <param name="begin_date">ტრანსპორტირების დაწყების თარიღი</param>
        /// <param name="waybill_number">ზედნადების ნომერი</param>
        /// <param name="errorText"></param>
        /// <returns>თუ წარმატებით გააქტიურდა მაშინ "1", თუ არა get_error_codes() ის დაბრუნებული ტექტსი, ან NULL</returns>
        public string send_waybill_transporter(int waybill_id, DateTime begin_date, out string waybill_number, out string errorText)
        {
            errorText = null;
            int send_waybill_transporter_result;
            string send_waybill_transporter_result_str = null;
            waybill_number = "";
            try
            {
                if (protocolType == "HTTP")
                {
                    send_waybill_transporter_result = wayBill_soapClient_HTTP.send_waybill_transporter(su, sp, waybill_id, begin_date, out waybill_number);
                }
                else
                {
                    send_waybill_transporter_result = wayBill_soapClient_HTTPS.send_waybill_transporter(su, sp, waybill_id, begin_date, out waybill_number);
                }
            }
            catch (Exception ex)
            {
                errorText = ex.Message + "! send_waybill_transporter()";
                return send_waybill_transporter_result_str;
            }

            if (send_waybill_transporter_result == 1)
            {
                return send_waybill_transporter_result.ToString();
            }
            else
            {
                string get_error_codes_result = null;
                get_error_codes_result = get_error_codes(send_waybill_transporter_result.ToString(), "", out errorText);
                if (errorText == null)
                {
                    send_waybill_transporter_result_str = get_error_codes_result;
                }
            }

            return send_waybill_transporter_result_str;
        }

        /// <summary>ზედნადების დახურვა</summary>
        /// <param name="waybill_id">ზედნადების ID</param>
        /// <param name="errorText"></param>
        /// <returns>1-დაიხურა; -1 არა; -101 სხვისი ზედნადებია და ვერ დახურავთ; -100 სერვისის მომხმარებელი ან პაროლი არასწორია.</returns>
        public int close_waybill(int waybill_id, out string errorText)
        {
            errorText = null;
            int close_waybill_result = -1;
            try
            {
                if (protocolType == "HTTP")
                {
                    close_waybill_result = wayBill_soapClient_HTTP.close_waybill(su, sp, waybill_id);
                }
                else
                {
                    close_waybill_result = wayBill_soapClient_HTTPS.close_waybill(su, sp, waybill_id);
                }

                if (close_waybill_result != 1)
                {
                    //if (close_waybill_result == -101)
                    //{
                    //    errorText = "სხვისი ზედნადებია და ვერ დახურავთ" + "! close_waybill()";
                    //    return close_waybill_result;
                    //}
                    //else
                    //{
                    string TYPE = "";
                    string get_error_codes_result = null;

                    get_error_codes_result = get_error_codes(close_waybill_result.ToString(), TYPE, out errorText);
                    if (get_error_codes_result != null)
                    {
                        errorText = get_error_codes_result;
                        return close_waybill_result;
                    }
                    //}
                }
            }
            catch (Exception ex)
            {
                errorText = ex.Message;
                return close_waybill_result;
            }

            return close_waybill_result;
        }

        /// <summary>ზედნადების დახურვა (delivery_date)</summary>
        /// <param name="delivery_date">მიწოდების თარიღი</param>
        /// <param name="waybill_id">ზედნადების ID</param>
        /// <param name="errorText"></param>
        /// <returns>1-დაიხურა; -1 არა; -101 სხვისი ზედნადებია და ვერ დახურავთ; -100 სერვისის მომხმარებელი ან პაროლი არასწორია.</returns>
        public int close_waybill_vd(DateTime delivery_date, int waybill_id, out string errorText)
        {
            errorText = null;
            int close_waybill_result = -1;
            try
            {
                if (protocolType == "HTTP")
                {
                    close_waybill_result = wayBill_soapClient_HTTP.close_waybill_vd(su, sp, delivery_date, waybill_id);
                }
                else
                {
                    close_waybill_result = wayBill_soapClient_HTTPS.close_waybill_vd(su, sp, delivery_date, waybill_id);
                }

                if (close_waybill_result != 1)
                {
                    //if (close_waybill_result == -101)
                    //{
                    //    errorText = "სხვისი ზედნადებია და ვერ დახურავთ" + "! close_waybill()";
                    //    return close_waybill_result;
                    //}
                    //else
                    //{
                    string TYPE = "";
                    string get_error_codes_result = null;

                    get_error_codes_result = get_error_codes(close_waybill_result.ToString(), TYPE, out errorText);
                    if (get_error_codes_result != null)
                    {
                        errorText = get_error_codes_result + "! close_waybill_vd()";
                        return close_waybill_result;
                    }
                    //}
                }
            }
            catch (Exception ex)
            {
                errorText = ex.Message + "! close_waybill_vd()";
                return close_waybill_result;
            }

            return close_waybill_result;
        }

        /// <summary>ზედნადების დახურვა (reception_info, receiver_info, delivery_date)</summary>
        /// <param name="waybill_id">ზედნადების ID</param>
        /// <param name="reception_info">მიმწოდებლის ინფორმაცია</param>
        /// <param name="receiver_info">მიმღების ინფორმაცია</param>
        /// <param name="delivery_date">მიწოდების თარიღი</param>
        /// <param name="errorText"></param>
        /// <returns>1-დაიხურა; -1 არა; -101 სხვისი ზედნადებია და ვერ დახურავთ; -100 სერვისის მომხმარებელი ან პაროლი არასწორია.</returns>
        public int close_waybill_transporter(int waybill_id, string reception_info, string receiver_info, DateTime delivery_date, out string errorText)
        {
            errorText = null;
            int close_waybill_result = -1;
            try
            {
                if (protocolType == "HTTP")
                {
                    close_waybill_result = wayBill_soapClient_HTTP.close_waybill_transporter(su, sp, waybill_id, reception_info, receiver_info, delivery_date);
                }
                else
                {
                    close_waybill_result = wayBill_soapClient_HTTPS.close_waybill_transporter(su, sp, waybill_id, reception_info, receiver_info, delivery_date);
                }

                if (close_waybill_result != 1)
                {
                    //if (close_waybill_result == -101)
                    //{
                    //    errorText = "სხვისი ზედნადებია და ვერ დახურავთ" + "! close_waybill()";
                    //    return close_waybill_result;
                    //}
                    //else
                    //{
                    string TYPE = "";
                    string get_error_codes_result = null;

                    get_error_codes_result = get_error_codes(close_waybill_result.ToString(), TYPE, out errorText);
                    if (get_error_codes_result != null)
                    {
                        errorText = get_error_codes_result + "! close_waybill_transporter()";
                        return close_waybill_result;
                    }
                    //}
                }
            }
            catch (Exception ex)
            {
                errorText = ex.Message + "! close_waybill_transporter()";
                return close_waybill_result;
            }

            return close_waybill_result;
        }

        /// <summary>ზედნადების წაშლა</summary>
        /// <param name="waybill_id">ზედნადების ID</param>
        /// <param name="errorText"></param>
        /// <returns>1-წაიშალა; -1 არა; -101 სხვისი ზედნადებია და ვერ წაშლით -100 სერვისის მომხმარებელი ან პაროლი არასწორია</returns>
        public int del_waybill(int waybill_id, out string errorText)
        {
            errorText = null;
            int del_waybill_result = -1;
            try
            {
                if (protocolType == "HTTP")
                {
                    del_waybill_result = wayBill_soapClient_HTTP.del_waybill(su, sp, waybill_id);
                }
                else
                {
                    del_waybill_result = wayBill_soapClient_HTTPS.del_waybill(su, sp, waybill_id);
                }
            }
            catch (Exception ex)
            {
                errorText = ex.Message + "! del_waybill()";
                return del_waybill_result;
            }

            return del_waybill_result;
        }

        /// <summary>ზედნადების გაუქმება</summary>
        /// <param name="waybill_id">ზედნადების ID</param>
        /// <param name="errorText"></param>
        /// <returns>1-გაუქმდა; -1 არა; -101 სხვისი ზედნადებია და ვერ გააუქმებთ -100 სერვისის მომხმარებელი ან პაროლი არასწორია</returns>
        public int ref_waybill(int waybill_id, out string errorText)
        {
            errorText = null;
            int ref_waybill_result = -1;
            try
            {
                if (protocolType == "HTTP")
                {
                    ref_waybill_result = wayBill_soapClient_HTTP.ref_waybill(su, sp, waybill_id);
                }
                else
                {
                    ref_waybill_result = wayBill_soapClient_HTTPS.ref_waybill(su, sp, waybill_id);
                }
            }
            catch (Exception ex)
            {
                errorText = ex.Message;
                return ref_waybill_result;
            }

            return ref_waybill_result;
        }

        /// <summary>ზედნადების გაუქმება (comment)</summary>
        /// <param name="waybill_id">ზედნადების ID</param>
        /// <param name="comment">კომენტარი</param>
        /// <param name="errorText"></param>
        /// <returns>1-გაუქმდა; -1 არა; -101 სხვისი ზედნადებია და ვერ გააუქმებთ -100 სერვისის მომხმარებელი ან პაროლი არასწორია</returns>
        public int ref_waybill_vd(int waybill_id, string comment, out string errorText)
        {
            errorText = null;
            int ref_waybill_result = -1;
            try
            {
                if (protocolType == "HTTP")
                {
                    ref_waybill_result = wayBill_soapClient_HTTP.ref_waybill_vd(su, sp, waybill_id, comment);
                }
                else
                {
                    ref_waybill_result = wayBill_soapClient_HTTPS.ref_waybill_vd(su, sp, waybill_id, comment);
                }
            }
            catch (Exception ex)
            {
                errorText = ex.Message;
                return ref_waybill_result;
            }

            return ref_waybill_result;
        }

        /// <summary>ზედნადების უარყოფა</summary>
        /// <param name="waybill_id">ზედნადების ID</param>
        /// <param name="errorText"></param>
        /// <returns>true თუ მდგომარეობა გახდა უარყოფილი, თუ არადა false (mara sul true-s abrunebs rac gavteste :))</returns>
        public bool reject_waybill(int waybill_id, out string errorText)
        {
            errorText = null;
            bool ref_waybill_result = false;
            try
            {
                if (protocolType == "HTTP")
                {
                    ref_waybill_result = wayBill_soapClient_HTTP.reject_waybill(su, sp, waybill_id);
                }
                else
                {
                    ref_waybill_result = wayBill_soapClient_HTTPS.reject_waybill(su, sp, waybill_id);
                }
            }
            catch (Exception ex)
            {
                errorText = ex.Message + "! reject_waybill()";
                return ref_waybill_result;
            }

            return ref_waybill_result;
        }

        /// <summary>მიღებული ზედნადების დადასტურება</summary>
        /// <param name="WBID"></param>
        /// <param name="errorText"></param>
        /// <returns></returns>
        public bool confirm_waybill(int WBID, out string errorText)
        {
            errorText = null;
            bool confirm_waybill_result = false;
            try
            {
                if (protocolType == "HTTP")
                {
                    confirm_waybill_result = wayBill_soapClient_HTTP.confirm_waybill(su, sp, WBID);
                }
                else
                {
                    confirm_waybill_result = wayBill_soapClient_HTTPS.confirm_waybill(su, sp, WBID);
                }
            }
            catch (Exception ex)
            {
                errorText = ex.Message + "! confirm_waybill()";
                return confirm_waybill_result;
            }

            return confirm_waybill_result;
        }

        //<--- სერვისის ელ. ზედნადების წარმოების მეთოდები


        //სასაქონლო კოდების აღრიცხვა --->

        /// <summary>შტრიხკოდის შენახვა (!)</summary>
        /// <param name="bar_code">სასაქონლო კოდი - შტრიხკოდი</param>
        /// <param name="goods_name">საქონლის სახელი</param>
        /// <param name="unit_id">ერთეულის ID</param>
        /// <param name="unit_txt">ერთეულის სახელი </param>
        /// <param name="a_id">აქციზის ID</param>
        /// <param name="errorText"></param>
        /// <returns>1-ოპერაცია შესრულდა; -1 არა;-100 სერვისის მომხმარებელი ან პაროლი არასწორია</returns>
        public int save_bar_code(string bar_code, string goods_name, int unit_id, string unit_txt, int a_id, out string errorText)
        {
            errorText = null;
            int save_bar_code_result = -1;
            try
            {
                if (protocolType == "HTTP")
                {
                    save_bar_code_result = wayBill_soapClient_HTTP.save_bar_code(su, sp, bar_code, goods_name, unit_id, unit_txt, a_id);
                }
                else
                {
                    save_bar_code_result = wayBill_soapClient_HTTPS.save_bar_code(su, sp, bar_code, goods_name, unit_id, unit_txt, a_id);
                }
            }
            catch (Exception ex)
            {
                errorText = ex.Message + "! save_bar_code()";
                return save_bar_code_result;
            }

            return save_bar_code_result;
        }

        /// <summary>შტრიხკოდის წაშლა</summary>
        /// <param name="bar_code">სასაქონლო კოდი - შტრიხკოდი</param>
        /// <param name="errorText"></param>
        /// <returns>1-ოპერაცია შესრულდა; -1 არა;-100 სერვისის მომხმარებელი ან პაროლი არასწორია</returns>
        public int delete_bar_code(string bar_code, out string errorText)
        {
            errorText = null;
            int delete_bar_code_result = -1;
            try
            {
                if (protocolType == "HTTP")
                {
                    delete_bar_code_result = wayBill_soapClient_HTTP.delete_bar_code(su, sp, bar_code);
                }
                else
                {
                    delete_bar_code_result = wayBill_soapClient_HTTPS.delete_bar_code(su, sp, bar_code);
                }
            }
            catch (Exception ex)
            {
                errorText = ex.Message + "! delete_bar_code()";
                return delete_bar_code_result;
            }

            return delete_bar_code_result;
        }

        /// <summary>შტრიკოდების სიის გამოტანა</summary>
        /// <param name="bar_code">სასაქონლო კოდი - შტრიხკოდი (თუ "" გადავეცით მაშინ სრული სია)</param>
        /// <param name="get_bar_codes_result_map">მიღებული შტრიხკოდების სია</param>
        /// <param name="errorText"></param>
        /// <returns>1-ოპერაცია შესრულდა; -1 არა;-100 სერვისის მომხმარებელი ან პაროლი არასწორია</returns>
        public int get_bar_codes(string bar_code, out Dictionary<string, HashSet<string>> get_bar_codes_result_map, out string errorText)
        {
            errorText = null;
            int get_bar_codes_result = -1;
            XmlNode bar_codes = null;
            get_bar_codes_result_map = null;
            try
            {
                if (protocolType == "HTTP")
                {
                    get_bar_codes_result = wayBill_soapClient_HTTP.get_bar_codes(su, sp, bar_code, out bar_codes);
                }
                else
                {
                    get_bar_codes_result = wayBill_soapClient_HTTPS.get_bar_codes(su, sp, bar_code, out bar_codes);
                }
            }
            catch (Exception ex)
            {
                errorText = ex.Message + "! get_bar_codes()";
                return get_bar_codes_result;
            }
            if (get_bar_codes_result == 1)
            {
                get_bar_codes_result_map = new Dictionary<string, HashSet<string>>();

                XmlNodeList itemNodes = bar_codes.SelectNodes("BAR_CODE");
                foreach (XmlNode itemNode in itemNodes)
                {
                    string CODE = (itemNode.SelectSingleNode("CODE") == null) ? "" : itemNode.SelectSingleNode("CODE").InnerText;
                    string NAME = (itemNode.SelectSingleNode("NAME") == null) ? "" : itemNode.SelectSingleNode("NAME").InnerText;
                    string UNIT_ID = (itemNode.SelectSingleNode("UNIT_ID") == null) ? "" : itemNode.SelectSingleNode("UNIT_ID").InnerText;
                    string UNIT_TXT = (itemNode.SelectSingleNode("UNIT_TXT") == null) ? "" : itemNode.SelectSingleNode("UNIT_TXT").InnerText;
                    get_bar_codes_result_map.Add(CODE, new HashSet<string>() { NAME, UNIT_ID, UNIT_TXT });
                }
            }

            return get_bar_codes_result;
        }

        //<--- სასაქონლო კოდების აღრიცხვა 


        //დისტრიბუციისათვის განკუთვნილი ავტომობილების აღრიცხვის მითოდი --->

        /// <summary>მანქანის ნომრის შენახვა</summary>
        /// <param name="car_number">მანქანის ნომერი</param>
        /// <param name="errorText"></param>
        /// <returns>1-ოპერაცია შესრულდა; -1 არა;-100 სერვისის მომხმარებელი ან პაროლი არასწორია</returns>
        public int save_car_numbers(string car_number, out string errorText)
        {
            errorText = null;
            int save_car_numbers_result = -1;
            try
            {
                if (protocolType == "HTTP")
                {
                    save_car_numbers_result = wayBill_soapClient_HTTP.save_car_numbers(su, sp, car_number);
                }
                else
                {
                    save_car_numbers_result = wayBill_soapClient_HTTPS.save_car_numbers(su, sp, car_number);
                }
            }
            catch (Exception ex)
            {
                errorText = ex.Message + "! save_car_numbers()";
                return save_car_numbers_result;
            }

            return save_car_numbers_result;
        }

        /// <summary>მანქანის ნომრის წაშლა</summary>
        /// <param name="car_number">მანქანის ნომერი</param>
        /// <param name="errorText"></param>
        /// <returns>1-ოპერაცია შესრულდა; -1 არა;-100 სერვისის მომხმარებელი ან პაროლი არასწორია</returns>
        public int delete_car_numbers(string car_number, out string errorText)
        {
            errorText = null;
            int delete_car_numbers_result = -1;
            try
            {
                if (protocolType == "HTTP")
                {
                    delete_car_numbers_result = wayBill_soapClient_HTTP.delete_car_numbers(su, sp, car_number);
                }
                else
                {
                    delete_car_numbers_result = wayBill_soapClient_HTTPS.delete_car_numbers(su, sp, car_number);
                }
            }
            catch (Exception ex)
            {
                errorText = ex.Message + "! delete_car_numbers()";
                return delete_car_numbers_result;
            }

            return delete_car_numbers_result;
        }

        /// <summary>მანქანის ნომრების სიის გამოტანა</summary>
        /// <param name="car_numbers_result_map">მიღებული მანქანის ნომრების სია</param>
        /// <param name="errorText"></param>
        /// <returns>1-ოპერაცია შესრულდა; -1 არა;-100 სერვისის მომხმარებელი ან პაროლი არასწორია</returns>
        public int get_car_numbers(out List<string> car_numbers_result_map, out string errorText)
        {
            errorText = null;
            int get_car_numbers_result = -1;
            XmlNode car_numbers = null;
            car_numbers_result_map = null;
            try
            {
                if (protocolType == "HTTP")
                {
                    get_car_numbers_result = wayBill_soapClient_HTTP.get_car_numbers(su, sp, out car_numbers);
                }
                else
                {
                    get_car_numbers_result = wayBill_soapClient_HTTPS.get_car_numbers(su, sp, out car_numbers);
                }
            }
            catch (Exception ex)
            {
                errorText = ex.Message + "! get_car_numbers()";
                return get_car_numbers_result;
            }
            if (get_car_numbers_result == 1)
            {
                car_numbers_result_map = new List<string>();

                XmlNodeList itemNodes = car_numbers.SelectNodes("CAR_NUMBER");
                foreach (XmlNode itemNode in itemNodes)
                {

                    car_numbers_result_map.Add((itemNode.SelectSingleNode("CAR_NUMBER") == null) ? "" : itemNode.SelectSingleNode("CAR_NUMBER").InnerText);
                }
            }

            return get_car_numbers_result;
        }

        //<--- დისტრიბუციისათვის განკუთვნილი ავტომობილების აღრიცხვის მითოდი


        //დამხმარე ფუქციები --->

        /// <summary>საიდენტიფიკაციო კოდით ან პირადი ნომრით სახელის გამოტანა</summary>
        /// <param name="tin">გადამხდელის საიდენტიფიკაციო ნომერი ან პირადი ნომერი</param>
        /// <param name="errorText"></param>
        /// <returns>მყიდველის სახელი, ან NULL</returns>
        public string get_name_from_tin(string tin, out string errorText)
        {
            errorText = null;
            string get_name_from_tin_result = null;
            try
            {
                if (protocolType == "HTTP")
                {
                    get_name_from_tin_result = wayBill_soapClient_HTTP.get_name_from_tin(su, sp, tin);
                }
                else
                {
                    get_name_from_tin_result = wayBill_soapClient_HTTPS.get_name_from_tin(su, sp, tin);
                }
            }
            catch (Exception ex)
            {
                errorText = ex.Message + "! get_name_from_tin()";
                return get_name_from_tin_result;
            }

            return get_name_from_tin_result;
        }

        /// <summary>შეცდომების კოდები</summary>
        /// <param name="STATUS">კოდი</param>
        /// <param name="TYPE">"" ყველა,  1 ძირითადი ზედნადების, 2 საქონლის ჩანაწერების, 3 ანგარიშფაქტურის გამოწერის შეცდომები</param>
        /// <param name="errorText"></param>
        /// <returns>შეცდომის ტექსტი ან NULL</returns>
        public string get_error_codes(string STATUS, string TYPE, out string errorText)
        {
            errorText = null;
            XmlNode get_error_codes_result = null;
            try
            {
                if (protocolType == "HTTP")
                {
                    get_error_codes_result = wayBill_soapClient_HTTP.get_error_codes(su, sp);
                }
                else
                {
                    get_error_codes_result = wayBill_soapClient_HTTPS.get_error_codes(su, sp);
                }
            }
            catch (Exception ex)
            {
                errorText = ex.Message + "! get_error_codes()";
                return null;
            }

            if (get_error_codes_result != null)
            {
                XmlNodeList itemNodes = get_error_codes_result.SelectNodes("ERROR_CODE");
                string TEXT = "";
                foreach (XmlNode itemNode in itemNodes)
                {
                    string ID = (itemNode.SelectSingleNode("ID") == null) ? "" : itemNode.SelectSingleNode("ID").InnerText;
                    string TYPE_RS = (itemNode.SelectSingleNode("TYPE") == null) ? null : itemNode.SelectSingleNode("TYPE").InnerText;
                    string TEXT_RS = (itemNode.SelectSingleNode("TEXT") == null) ? "" : itemNode.SelectSingleNode("TEXT").InnerText;
                    if (ID == STATUS && TYPE == "" || ID == STATUS && TYPE == TYPE_RS)
                    {
                        TEXT = TEXT + "\n" + TEXT_RS;
                    }
                }
                return TEXT;
            }

            return null;
        }

        /// <summary>IP-ის გაგება</summary>
        /// <param name="errorText"></param>
        /// <returns>IP ან NULL</returns>
        public string what_is_my_ip(out string errorText)
        {
            errorText = null;
            string my_ip = null;
            try
            {
                if (protocolType == "HTTP")
                {
                    my_ip = wayBill_soapClient_HTTP.what_is_my_ip();
                }
                else
                {
                    my_ip = wayBill_soapClient_HTTPS.what_is_my_ip();
                }
            }
            catch (Exception ex)
            {
                errorText = ex.Message + "! what_is_my_ip()";
            }
            return my_ip;
        }

        /// <summary>არის თუ არა დღგ-ის გადამხდელი</summary>
        /// <param name="tin">გადამხდელის საიდენტიფიკაციო ნომერი ან პირადი ნომერი</param>
        /// <param name="errorText"></param>
        /// <returns></returns>
        public bool is_vat_payer_tin(string tin, out string errorText)
        {
            errorText = null;
            bool is_vat_payer_tin_result = false;
            try
            {
                if (protocolType == "HTTP")
                {
                    is_vat_payer_tin_result = wayBill_soapClient_HTTP.is_vat_payer_tin(su, sp, tin);
                }
                else
                {
                    is_vat_payer_tin_result = wayBill_soapClient_HTTPS.is_vat_payer_tin(su, sp, tin);
                }
            }
            catch (Exception ex)
            {
                errorText = ex.Message + "! is_vat_payer_tin()";
                return is_vat_payer_tin_result;
            }

            return is_vat_payer_tin_result;
        }

        /// <summary>
        /// get_waybill_goods_list
        /// </summary>
        /// <param name="errorText"></param>
        /// <returns></returns>
        public DataTable get_waybill_goods_list(DateTime begin_date_s, DateTime begin_date_e, string itypes, string buyer_tin, string statuses, string car_number, string waybill_number, out string errorText)
        {
            errorText = null;
            XmlNode XmlData = null;
            Dictionary<string, HashSet<string>> goods_lift_map = null;

            DataTable RSGoodsTable = new DataTable();

            try
            {
                if (protocolType == "HTTP")
                {
                    XmlData = wayBill_soapClient_HTTP.get_waybill_goods_list(su, sp, itypes, buyer_tin, statuses, car_number, begin_date_s, begin_date_e, null, null, null, null, null, null, waybill_number, null, null, null, null);
                }
                else
                {
                    XmlData = wayBill_soapClient_HTTPS.get_waybill_goods_list(su, sp, itypes, buyer_tin, statuses, car_number, begin_date_s, begin_date_e, null, null, null, null, null, null, waybill_number, null, null, null, null);
                }
            }
            catch
            {
                //errorText = ex.Message + "! is_vat_payer_tin()";
                //return is_vat_payer_tin_result;
            }

            if (XmlData != null)
            {
                goods_lift_map = new Dictionary<string, HashSet<string>>();
                StringReader sr = new StringReader("<root>" + XmlData.InnerXml + "</root>");

                DataSet RSDataSet = new DataSet();
                RSDataSet.ReadXml(sr);

                try
                {
                    RSGoodsTable = RSDataSet.Tables[0];
                }
                catch
                { }

                //XmlNodeList itemNodes = XmlData.SelectNodes("ServiceUser");
                //foreach (XmlNode itemNode in itemNodes)
                //{
                //    string ID = (itemNode.SelectSingleNode("ID") == null) ? "" : itemNode.SelectSingleNode("ID").InnerText;
                //    string USER_NAME = (itemNode.SelectSingleNode("USER_NAME") == null) ? "" : itemNode.SelectSingleNode("USER_NAME").InnerText;
                //    string UN_ID = (itemNode.SelectSingleNode("UN_ID") == null) ? "" : itemNode.SelectSingleNode("UN_ID").InnerText;
                //    string IP = (itemNode.SelectSingleNode("IP") == null) ? "" : itemNode.SelectSingleNode("IP").InnerText;
                //    string NAME = (itemNode.SelectSingleNode("NAME") == null) ? "" : itemNode.SelectSingleNode("NAME").InnerText;
                //    goods_lift_map.Add(ID, new HashSet<string>() { USER_NAME, UN_ID, IP, NAME });
                //}
            }


            List<string> ColumnsList = new List<string>();
            ColumnsList.Add("WAYBILL_NUMBER");
            ColumnsList.Add("FULL_AMOUNT");
            ColumnsList.Add("STATUS");
            ColumnsList.Add("SELLER_TIN");
            ColumnsList.Add("SELLER_NAME");
            ColumnsList.Add("BEGIN_DATE");
            ColumnsList.Add("TYPE");
            ColumnsList.Add("START_ADDRESS");
            ColumnsList.Add("END_ADDRESS");
            ColumnsList.Add("DELIVERY_DATE");
            ColumnsList.Add("ACTIVATE_DATE");
            ColumnsList.Add("CAR_NUMBER");
            ColumnsList.Add("DRIVER_TIN");
            ColumnsList.Add("TRANSPORT_COAST");
            ColumnsList.Add("W_NAME");
            ColumnsList.Add("BAR_CODE");
            ColumnsList.Add("AMOUNT");
            ColumnsList.Add("Quantity");

            foreach (string ColName in ColumnsList)
            {
                if (RSGoodsTable.Columns.Contains(ColName) == false)
                {
                    RSGoodsTable.Columns.Add(ColName);
                }

            }

            return RSGoodsTable;

        }


        //<--- დამხმარე ფუქციები
    }
}