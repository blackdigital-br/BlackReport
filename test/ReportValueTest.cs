namespace BlackDigital.Report.Test
{
    public class ReportValueTest
    {
        [Fact]
        public void SingleReportValue()
        {
            var value = new SingleReportValue("Hello World");

            Assert.Null(value.GetValue());

            Assert.True(value.NextRow());
            Assert.True(value.NextColumn());

            Assert.Equal("Hello World", value.GetValue());

            Assert.False(value.NextColumn());
            Assert.Null(value.GetValue());

            Assert.False(value.NextRow());
            Assert.Null(value.GetValue());
        }
    }
}