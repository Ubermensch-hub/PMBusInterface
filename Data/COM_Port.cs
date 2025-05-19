using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PMBusInterface.Data
{
    public class COM_Port
    {
        public string Name { get; set; }
        public string Description { get; set; }

        public override string ToString() => Name;
       
    }
}
