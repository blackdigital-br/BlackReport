using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BlackDigital.Report
{
    public record ValueFormatter
    {
        public object? Value { get; init; }

        public Type? ValueType { get; init; }

        public IFormatProvider? FormatProvider { get; init; }

        public string? Format { get; init; }
    }
}
