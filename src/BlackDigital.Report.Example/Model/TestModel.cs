using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
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
            Id = DateTime.Now.Ticks;
            Name = name;
            Number = number;

            ObjDate = objDate;
            Time = time;

#if NET6_0_OR_GREATER

            ObjDate2 = objDate2;
            Time2 = time2;
#endif

        }


        [Display(Order = 1, AutoGenerateField = false)]
        public long Id { get; set; }

        [Display(Order = 1)]
        public string Name { get; set; }

        [Display(Order = 2)]
        public double Number { get; set; }

        [Display(Order = 4)]
        public DateTime ObjDate { get; set; }

        [Display(Order = 3)]
        public TimeSpan Time { get; set; }

#if NET6_0_OR_GREATER

        public DateOnly ObjDate2 { get; set; }

        public TimeOnly Time2 { get; set; }

#endif
    }
}
