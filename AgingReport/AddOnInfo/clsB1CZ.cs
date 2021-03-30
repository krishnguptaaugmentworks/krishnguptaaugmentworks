using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data.SqlClient;
using System.Runtime.InteropServices;
using SAPbobsCOM;
using System.Xml;
using System.Windows.Forms;

namespace AgingReport
{
    public partial class AddOnInfo
    {
        public void B1CZ_ItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED && pVal.BeforeAction == false && pVal.ItemUID == "FTCW")
                {
                    FTCW_AfterAddBtnPressed(FormUID, ref pVal, out BubbleEvent);
                }
            }
            catch (SqlException e)
            {
                __app.MessageBox(e.Message, 1, "Ok", "", "");
                BubbleEvent = false;
            }
            catch (COMException e)
            {
                __app.MessageBox(e.Message, 1, "Ok", "", "");
                BubbleEvent = false;
            }
            catch (Exception e)
            {
                __app.MessageBox(e.Message, 1, "Ok", "", "");
                BubbleEvent = false;
            }
        }

        private void FTCW_AfterAddBtnPressed(string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;

            __Form = __app.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount);
            oMatrix = __Form.Items.Item("Item_1").Specific;
            oMatrix.AutoResizeColumns();
        }
    }
}
