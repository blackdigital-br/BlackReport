using System;

namespace BlackDigital.Report
{
    public class ValueFormatter
    {
        public IFormatProvider? FormatProvider { get; init; }

        public string? Format { get; init; }

        public string? TypeName { get; init; }
    }
}
