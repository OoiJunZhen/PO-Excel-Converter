using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PO_Excel.Model
{
    public class DataList
    {
        public string PRProjectCode { get; set; }
        public string PRNo { get; set; }
        public string? PRApprovedOn { get; set; }
        public string PRMaterialCode { get; set; }
        public string PRQty { get; set; }
        public string POProjectCode { get; set; }
        public string POMaterialCode { get; set; }
        public string? PONo { get; set; }
        public string? POApprovedOn { get; set; }
        public string POQty { get; set; }
        public string? Supplier { get; set; }
        public string? ReceivedQty { get; set; }
        public string? ETA { get; set;}

    }
}
