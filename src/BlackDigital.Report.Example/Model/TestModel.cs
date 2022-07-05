using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BlackDigital.Report.Example.Model
{
    public class TestModel
    {
        public TestModel(string name, double number, DateTime objDate, TimeSpan time)
        {
            Name = name;
            Number = number;

            ObjDate = objDate;
            Time = time;            
        }

        public string Name { get; set; }

        public double Number { get; set; }

        public DateTime ObjDate { get; set; }

        public TimeSpan Time { get; set; }
    }
}
