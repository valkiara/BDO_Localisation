using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BDO_Localisation_AddOn
{
	static partial class CashFlowLineItem
	{
		public static void uiApp_MenuEvent(ref SAPbouiCOM.MenuEvent pVal, out bool BubbleEvent, out string errorText)
		{
			errorText = null;
			BubbleEvent = true;

			try
			{
				SAPbouiCOM.Form oForm = Program.uiApp.Forms.ActiveForm;
				SAPbouiCOM.DBDataSource oDBDataSource = oForm.DataSources.DBDataSources.Item(0);
				string code = oDBDataSource.GetValue("CFWId", 0).Trim();

				if (checkRemoving(oForm, out errorText))
				{
					Program.uiApp.StatusBar.SetSystemMessage(BDOSResources.getTranslate("CannotDeleteACashFlowLineItemAssignedIn") + BDOSResources.getTranslate("InternetBankingIntegrationServicesRules"));


					BubbleEvent = false;
				}
			}
			catch (Exception ex)
			{
				Program.uiApp.MessageBox(ex.ToString(), 1, "", "");
			}
		}

		public static bool checkRemoving(SAPbouiCOM.Form oForm, out string errorText)
		{
			errorText = null;

			SAPbouiCOM.DBDataSource oDBDataSource = oForm.DataSources.DBDataSources.Item(0);
			string code = oDBDataSource.GetValue("CFWId", 0).Trim();
			Dictionary<string, string> listTables = new Dictionary<string, string>();
			listTables.Add("@BDOSINTR", "U_CFWId");
			return CommonFunctions.codeIsUsed(listTables, code);
		}
	}
}
