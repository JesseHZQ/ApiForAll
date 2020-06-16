using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace ApiForControlTower.Models
{
    public class BoardLocation
    {
        public string PN { get; set; }
        public string Description { get; set; }
        public string Type { get; set; }
        public List<SupermarketBoard> SupermarketList { get; set; }
        public int SupermarketListQTY { get; set; }
        public List<VerifyBoard> VerifyList { get; set; }
        public int VerifyListQTY { get; set; }
        public List<TestBoard> TestList { get; set; }
        public int TestListQTY { get; set; }
        public List<BurnInBoard> BurnInList { get; set; }
        public int BurnInListQTY { get; set; }
    }

    public class BurnInBoard
    {
        public string PN { get; set; }
        public string SN { get; set; }
    }

    public class TestBoard
    {
        public string PN { get; set; }
        public string SN { get; set; }
        public string TesterID { get; set; }
        public string Slot { get; set; }
    }

    public class VerifyBoard
    {
        public string PN { get; set; }
        public string SN { get; set; }
    }
    public class SupermarketBoard
    {
        public string PN { get; set; }
        public string SN { get; set; }
        public string Location { get; set; }
    }


    public class RackDetail
    {
        public string PN { get; set; }
        public int ActualQTY { get; set; }
        public string RackName { get; set; }
        public string PNList { get; set; }
        public string SNList { get; set; }
        public string SlotList { get; set; }
    }

    public class TestSystem
    {
        public string System { get; set; }
        public string BayNO { get; set; }
    }
}