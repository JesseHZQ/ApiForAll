using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web.Http;

namespace ApiForFCTKBNew.Controllers
{
    public class RefreshController : ApiController
    {
        [HttpGet]
        public string Refresh()
        {
            string msg = "";
            SlotPlanController slotplan = new SlotPlanController();
            msg += slotplan.UpdateSlotPlan();
            SlotPOController slotpo = new SlotPOController();
            msg += slotpo.UpdateSlotPO();
            //SlotCORCController slotcorc = new SlotCORCController();
            //msg += slotcorc.UpdateSlotCORC();
            //SlotShortageController slotshortage = new SlotShortageController();
            //msg += slotshortage.UpdateSlotShortage();
            //SlotConfigController slotconfig = new SlotConfigController();
            //msg += slotconfig.UpdateSlotConfig();
            //MinioneController minione = new MinioneController();
            //msg += minione.UpdateSlotMinioneStatus();
            //SlotStatusController slotstatus = new SlotStatusController();
            //msg += slotstatus.UpdateSlotStatus();
            //ShippingController shipping = new ShippingController();
            //msg += shipping.UpdateShippingStatus();
            //FCCPController fccp = new FCCPController();
            //msg += fccp.UpdateFCCP();
            return msg;
        }
    }
}
