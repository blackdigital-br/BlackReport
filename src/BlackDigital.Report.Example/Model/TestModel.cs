using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BlackDigital.Report.Example.Model
{
    public class TestModel
    {
        public TestModel(string name, 
                        double number, 
                        DateTime objDate, 
                        TimeSpan time
#if NET6_0_OR_GREATER
                        ,
                        DateOnly objDate2,
                        TimeOnly time2
#endif
            )
        {
            Name = name;
            Number = number;

            ObjDate = objDate;
            Time = time;

#if NET6_0_OR_GREATER

            ObjDate2 = objDate2;
            Time2 = time2;
#endif

        }

        public string Name { get; set; }

        public double Number { get; set; }

        public DateTime ObjDate { get; set; }

        public TimeSpan Time { get; set; }

#if NET6_0_OR_GREATER

        public DateOnly ObjDate2 { get; set; }

        public TimeOnly Time2 { get; set; }

#endif
    }
}
