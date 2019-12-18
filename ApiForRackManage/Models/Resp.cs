using ApiForRackManage.DataContext;
using System.Collections.Generic;

namespace ApiForRackManage.Controllers
{
    public class Resp
    {
        public int Code { get; set; }
        public List<Instrument> DataUF { get; set; }
        public List<J750_Board> Data750 { get; set; }
        public string Message { get; set; }
    }
}
