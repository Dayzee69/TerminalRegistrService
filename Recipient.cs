
using System.Collections.Generic;

namespace TerminalRegistrService
{
    class Recipient
    {

        public string recipientID { get; set; }
        public string descriptionName { get; set; }
        public string description { get; set; }
        public string internalName { get; set; }
        public bool exluse { get; set; }
        public string name { get; set; }
        public string email = "";
        public string queryColumns = "";
        public string[] headersColumns { get; set;  }
        public Dictionary<string, string> columns = new Dictionary<string, string>();

        public Recipient(string recipientID, string internalName) 
        {
            this.recipientID = recipientID;
            this.internalName = internalName;
        }

    }
}
