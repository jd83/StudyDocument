using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PrepareForWork
{
    public class OrderTransferRequestTransaction
    {
        public int RequestID { get; set; }
        public int SONumber { get; set; }
        public string FromWH { get; set; }
        public string ToWH { get; set; }
        public OrderTransferRequestPhase Phase { get; set; }
        public string ExceptionMessage { get; set; }

    }
    public enum OrderTransferRequestPhase
    {
        Hold,Q4S,Transferred
    }
}
