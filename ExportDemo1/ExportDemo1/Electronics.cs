using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExportDemo1
{
    public class Electronics
    {
        public int Id { get; set; }
        public DateTime Data { get; set; }
        public string Name { get; set; }

        public static List<Electronics> GetData()
        {
            return new List<Electronics>(){
                
            new Electronics() { Id = 1, Data = DateTime.Now.Date.AddDays(-3), Name = "TV" },
            new Electronics() { Id = 2, Data = DateTime.Now.Date.AddDays(-1), Name = "PC" },
            new Electronics() { Id = 3, Data = new DateTime(2003,11,5), Name = "Camera" }
            };

        }
    }
}
