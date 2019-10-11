using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace ApiForControlTower.Models
{
    public class User
    {
        public int UserId { get; set; }
        public string UserName { get; set; }
        public string Password { get; set; }
        public string Department { get; set; }
        public int LevelNum { get; set; }
        public string Shift { get; set; }
        public int IsLeave { get; set; }
        public int IsLock { get; set; }
        public string Skills { get; set; }
        public string Phone { get; set; }
        public string Email { get; set; }
        public int BackUpId { get; set; }
        public int IsDel { get; set; }
        public string WeChatName { get; set; }
    }
}