using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ProInterfaceFracttal
{
    public class Data
    {
        public int id { get; set; }
        public string warehouses_source_description { get; set; }
        public string code { get; set; }
        public string folio_source { get; set; }
        public string date { get; set; }
        public string document { get; set; }
        public string movements_states_description { get; set; }
        public string costs_center_description { get; set; }
        public string responsible_description { get; set; }
        public string create_by_description { get; set; }
        public string date_create { get; set; }
        public string number_details { get; set; }
        public string currency_description { get; set; }
        public string currency_symbol { get; set; }
        public List<list_items> list_items { get; set; }
    }
}
