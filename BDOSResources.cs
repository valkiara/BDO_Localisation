using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Reflection;
using System.Resources;
using System.Text;
using System.Threading;
using System.Windows.Forms;

namespace BDO_Localisation_AddOn
{
    static partial class BDOSResources
    {
        public static void initResource(int language, out CultureInfo cultureInfo, out ResourceManager resourceManager, out string errorText)
        {
            errorText = null;
            //cultureInfo = CultureInfo.GetCultureInfo("en-GB");
            cultureInfo = CultureInfo.CreateSpecificCulture("en");
            switch (language)
            {
                case 8: cultureInfo = CultureInfo.GetCultureInfo("en");
                    break;
                case 3: cultureInfo = CultureInfo.GetCultureInfo("ka");
                    break;
                case 100007: cultureInfo = CultureInfo.GetCultureInfo("ka");
                    break;
                case 24: cultureInfo = CultureInfo.GetCultureInfo("ru");
                    break;
            }

            resourceManager = new ResourceManager("BDO_Localisation_AddOn.Resource.Res", Assembly.GetExecutingAssembly());
            //resourceManager = ResourceManager.CreateFileBasedResourceManager("resource", "Resources", null);
        }

        public static string getTranslate(string key)
        {
            try
            {
                string translate = Program.resourceManager.GetString(key, Program.cultureInfo);

                if (translate == "" || translate == null)
                {
                    translate = key;
                }
                else if (translate == "space")
                {
                    return "";
                }
                return translate;
            }
            catch
            {
                if (char.IsLower(key[0]))
                {
                    key = firstLetterToUpper(key);
                }
                else if (char.IsUpper(key[0]))
                {
                    key = firstLetterToLower(key);
                }

                try
                {
                    string translate = Program.resourceManager.GetString(key, Program.cultureInfo);

                    if (translate == "" || translate == null)
                    {
                        translate = key;
                    }
                    else if (translate == "space")
                    {
                        return "";
                    }
                    return translate;
                }
                catch
                {
                    return key;
                }
            }
        }

        public static string firstLetterToUpper(string str)
        {
            if (str == null)
                return null;

            if (str.Length > 1)
                return char.ToUpper(str[0]) + str.Substring(1);

            return str.ToUpper();
        }

        public static string firstLetterToLower(string str)
        {
            if (str == null)
                return null;

            if (str.Length > 1)
                return char.ToLower(str[0]) + str.Substring(1);

            return str.ToLower();
        }
    }
}
