using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web.Http;

namespace ApiForCommonTools.Controllers
{
    public class ExcelController : ApiController
    {
        [HttpPost]
        public ExcelResp GenerateExcel(AcceptList list)
        {
            ExcelResp resp = new ExcelResp();
            return resp;
        }
    }

    public class ExcelResp
    {
    }

    public class AcceptList
    { 
        //泛型类，声明T变量类型
        public List<object> List { get; set; }//未定义具体类型的泛型集合
    }
}
