using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace BDO_Localisation_AddOn
{
    static partial class NumberToGeorgianTextConverter
    {
        private static string[] digit100_900 = new string[] { "ას", "ორას", "სამას", "ოთხას", "ხუთას", "ექვსას", "შვიდას", "რვაას", "ცხრაას" };
        private static string[] digit1000_ = new string[] { "ათას", "მილიონ", "მილიარდ", "ტრილიონ" };
        private static string[] digits1_19 = new string[] { 
           "ნული", "ერთი", "ორი", "სამი", "ოთხი", "ხუთი", "ექვსი", "შვიდი", "რვა", "ცხრა", "ათი", "თერთმეტი", "თორმეტი", "ცამეტი", "თოთხმეტი", "თხუთმეტი", 
           "თექვსმეტი", "ჩვიდმეტი", "თვრამეტი", "ცხრამეტი"
        };
        private static string[] digits20_80_long = new string[] { "ოცი", "ორმოცი", "სამოცი", "ოთხმოცი" };
        private static string[] digits20_80_short = new string[] { "ოცდა", "ორმოცდა", "სამოცდა", "ოთხმოცდა" };
        private static string minusSign = "მინუს";

        public static string Convert(decimal number, bool asCurrency, bool truncate, string currency, string currencyChange)
        {
            if (number > 1000000000000M)
            {
                throw new ApplicationException("ძალიან დიდი თანხა! " + number.ToString());
            }
            string text = "";
            if (number < 0M)
            {
                text = minusSign + " ";
                number -= 0M;
            }
            string text2 = null;
            long num = System.Convert.ToInt64(decimal.Truncate(number));
            long num2 = System.Convert.ToInt64((decimal)((number - decimal.Truncate(number)) * 100M));
            if (!truncate)
            {
                text2 = " ";
                if (asCurrency)
                {
                    text2 = ((text2 + currency) + " და " + ((num2.ToString().Length <= 1) ? ("0" + num2.ToString()) : num2.ToString())) + " " + currencyChange;
                }
                else
                {
                    text2 = (text2 + "მთელი  და " + ((num2.ToString().Length <= 1) ? ("0" + num2.ToString()) : num2.ToString())) + " " + "მეასედი";
                }
            }
            string text3 = GetTextUpper100(num);
            num /= (long)0x3e8;
            int index = 0;
            while (num != 0)
            {
                if (GetTextUpper100(num) != "")
                {
                    text3 = GetTextUpper100(num) + " " + digit1000_[index] + (((num % ((long)0x3e8)) == 0) ? "ი" : "") + " " + text3;
                }
                index++;
                num /= (long)0x3e8;
            }
            string text4 = text3;
            if ((text3.Substring(text3.Length - 3, 3) != "რვა") && (text3.Substring(text3.Length - 3, 3) != "ხრა"))
            {
                text4 = (text3[text3.Length - 1] != 'ი') ? (text3.Remove(text3.Length - 1, 1) + "ი") : text3;
            }
            return (text + text4 + text2);
        }

        private static string GetTextUnder20(long p_Value)
        {
            int length = p_Value.ToString().Length;
            string text = p_Value.ToString().Substring(length - 2, 2);
            long num2 = System.Convert.ToInt64(text) / ((long)10);
            long num3 = System.Convert.ToInt64(text) % ((long)10);
            string text2 = "";
            if (System.Convert.ToInt64(text) <= 0x13)
            {
                return digits1_19[(int)((IntPtr)System.Convert.ToInt64(text))];
            }
            if (num2 >= 4)
            {
                text2 = digits1_19[(int)((IntPtr)(num2 / ((long)2)))].Remove(digits1_19[(int)((IntPtr)(num2 / ((long)2)))].Length - 1, 1);
                if (text2.IndexOf("მ") == -1)
                {
                    text2 = text2 + "მ";
                }
            }
            return (text2 + "ოც" + (((System.Convert.ToInt64(text) % ((long)20)) == 0) ? "ი" : "და") + ((num3 == 0) ? (((System.Convert.ToInt64(text) % ((long)20)) != 0) ? "ათი" : "") : digits1_19[(int)((IntPtr)(System.Convert.ToInt64(text) - ((System.Convert.ToInt64(text) / ((long)20)) * 20)))]));
        }

        private static string GetTextUpper100(long p_Value)
        {
            int length = p_Value.ToString().Length;
            if (length <= 1)
            {
                return digits1_19[(int)((IntPtr)p_Value)];
            }
            if (length == 2)
            {
                return GetTextUnder20(p_Value);
            }
            string text = p_Value.ToString().Substring(length - 3, 3);
            if (System.Convert.ToInt64(text) == 0)
            {
                return "";
            }
            if (System.Convert.ToInt64(text) <= 0x13)
            {
                return digits1_19[(int)((IntPtr)System.Convert.ToInt64(text))];
            }
            if (System.Convert.ToInt64(text) <= 0x63)
            {
                return GetTextUnder20(System.Convert.ToInt64(text));
            }
            if (text == "000")
            {
                return "";
            }
            long num2 = System.Convert.ToInt64(text) / ((long)100);
            long num3 = System.Convert.ToInt64(text) % ((long)100);
            string text2 = ((num2 >= 2) ? digits1_19[(int)((IntPtr)num2)].Remove(digits1_19[(int)((IntPtr)num2)].Length - 1, 1) : "") + "ას";
            if ((num3 == 0) && (num2 != 0))
            {
                return (text2 + "ი");
            }
            return (((num2 >= 2) ? digits1_19[(int)((IntPtr)num2)].Remove(digits1_19[(int)((IntPtr)num2)].Length - 1, 1) : "") + "ას" + GetTextUnder20(p_Value));
        }
    }
}
