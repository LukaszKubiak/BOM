using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BOM
{
    class Codes
    {
        public string oldCode { get; set; }
        public string newCode { get; set; }

        public Codes(string oldCode,string newCode)
        {
            this.oldCode = oldCode;
            this.newCode = newCode;
        }
    }
}
