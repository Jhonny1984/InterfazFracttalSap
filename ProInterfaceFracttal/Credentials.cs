using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ProInterfaceFracttal
{
    public class Credentials
    {
        public string Id { get; set; }
        public string User { get; set; }
        public string Key { get; set; }
        public bool IsValid { get; }
    }
}
