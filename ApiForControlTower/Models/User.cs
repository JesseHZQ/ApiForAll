﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace ApiForControlTower.Models
{
    public class User
    {
        public string UserID { get; set; }
        public string UserName { get; set; }
        public string Password { get; set; }
        public int LevelNum { get; set; }
        public string Shift { get; set; }
        public string Skills { get; set; }
        public string Phone { get; set; }
        public string Email { get; set; }
        public int BackUpId { get; set; }
    }
}