using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ProInterfaceFracttal
{
    public class RootObject
    {

        public List<object> cookies { get; set; }
        public List<Requis> Requis { get; set; }
        public string connectorGuid { get; set; }
        public string connectorVersionGuid { get; set; }
        public string pageUrl { get; set; }
        public int offset { get; set; }
    }
}
