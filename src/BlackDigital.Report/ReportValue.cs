
namespace BlackDigital.Report
{
    internal abstract class ReportValue
    {
        internal ReportValue()
        {
        }
        
        public bool Processed { get; protected set; } = false;

        internal abstract bool NextRow();

        internal abstract bool NextColumn();

        internal abstract object? GetValue();
    }
}
