using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TerminalRegistrService
{
    class Payment : Recipient
    {
        public string paymentDate { get; set; }
        public string paymentInformationID { get; set; }
        public double amount { get; set; }
        public double comission { get; set; }
        public double fee { get; set; }
        public string session { get; set; }


        public Payment(string recipientID, string internalName, string query)
            : base(recipientID, internalName)
        {

            this.queryColumns = query;
        }
    }
}
