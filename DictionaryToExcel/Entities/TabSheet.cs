using System;
using System.Collections.Generic;
using System.Text;

namespace Ougha.Entities
{
    public class TabSheet
    {
        public string Name { get; set; }
        public virtual List<Dictionary<string,string>> Properties { get; set; }

        public TabSheet()
        {
            Properties = new List<Dictionary<string,string>>();
        }
    }
}
