using System.ComponentModel.DataAnnotations;

namespace BlackDigital.Report.Example.Model
{
    public class SimpleModel
    {
        public SimpleModel(string name, double number)
        {
            Name = name;
            Number = number;
        }
        [Display(Order = 1)]
        public string Name { get; set; }

        [Display(Order = 2)]
        public double Number { get; set; }
    }
}
