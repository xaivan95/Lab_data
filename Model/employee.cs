using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lab_data.Model
{
    public class employee
    {
        public employee() { }

        public employee(int id) {
            FIO = "";
            gender = "М";
            childre = 0;
            Family = "Не замужем (не женат)";
            degree = "-";
            title = "-";
            Post_id = id;
        }
        [Key]
        public int ID { get; set; }

        public string? FIO { get; set; }
        public string? gender { get; set; }
        public int? Age { get; set; }
        public string? Family { get; set; }
        
        public int? childre { get; set; }
        public int? Post_id { get; set; }
        public string? degree { get; set; }
        public string? title { get; set; }

    }
}
