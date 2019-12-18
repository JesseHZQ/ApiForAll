using System.Collections.Generic;

namespace ApiForRackManage.Controllers
{
    public class RespBI
    {
        public int Code { get; set; }
        public List<BI> List { get; set; }
        public string Message { get; set; }
    }
}
