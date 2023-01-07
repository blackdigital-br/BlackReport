using System;

namespace BlackDigital.Report
{
    public record ValueFormatter
    {
        public IFormatProvider? FormatProvider { get; init; }

        public string? Format { get; init; }
    }
}
