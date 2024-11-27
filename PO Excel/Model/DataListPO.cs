using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PO_Excel.Model
{
    public class DataListPO
    {
        public string POProjectCode { get; set; }
        public string POMaterialCode { get; set; }
        public string POQty { get; set; }
        public string PONo { get; set; }
        public string POApprovedOn { get; set; }
        public string ReceivedQty { get; set; }
    }
}
