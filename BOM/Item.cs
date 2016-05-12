using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BOM
{
    class Item
    {
        public string Line { get; set; }
        public string ItemCode { get; set; }
        public string ItemDesc { get; set; }
        public List<Child> Children { get; set; }
        public Item()
        {
            Children = new List<Child>();
        }

    }
}
