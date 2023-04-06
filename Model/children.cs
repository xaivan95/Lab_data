using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lab_data.Model
{
    public class children
    {
        public children() { }
        public children(int id) {
            FIO = "";
            birthdate = "";
            employee_id = id;
        }

        public int ID { get; set; }
        public string? FIO { get; set; }
        public string? birthdate { get; set; }
        public int employee_id { get; set; }
    }
}
