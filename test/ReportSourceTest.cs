using BlackDigital.Report.Example.Model;

namespace BlackDigital.Report.Test
{
    public class ReportSourceTest
    {
        [Fact]
        public void SingleReportValue()
        {
            var value = new SingleReportSource("Hello World");

            Assert.Null(value.GetValue());

            Assert.True(value.NextRow());
            Assert.True(value.NextColumn());

            Assert.Equal("Hello World", value.GetValue());

            Assert.False(value.NextColumn());
            Assert.Null(value.GetValue());

            Assert.False(value.NextRow());
            Assert.Null(value.GetValue());
        }

        [Fact]
        public void SingleReportValueWithNull()
        {
            var value = new SingleReportSource(null);

            Assert.Null(value.GetValue());

            Assert.True(value.NextRow());
            Assert.True(value.NextColumn());

            Assert.Null(value.GetValue());

            Assert.False(value.NextColumn());
            Assert.Null(value.GetValue());

            Assert.False(value.NextRow());
            Assert.Null(value.GetValue());
        }

        [Fact]
        public void EnumerableReportValue()
        {
            var value = new EnumerableReportSource(new[]
            {
                new[] { "Hello", "World" },
                new[] { "Hello2", "World2" }
            });

            Assert.Null(value.GetValue());

            Assert.True(value.NextRow());
            Assert.True(value.NextColumn());

            Assert.Equal("Hello", value.GetValue());

            Assert.True(value.NextColumn());
            Assert.Equal("World", value.GetValue());

            Assert.False(value.NextColumn());
            Assert.Null(value.GetValue());

            Assert.True(value.NextRow());
            Assert.Null(value.GetValue());

            Assert.True(value.NextColumn());
            Assert.Equal("Hello2", value.GetValue());

            Assert.True(value.NextColumn());
            Assert.Equal("World2", value.GetValue());

            Assert.False(value.NextColumn());
            Assert.False(value.NextRow());

            Assert.Null(value.GetValue());
        }

        [Fact]
        public void ModelReportSource()
        {
            var value = new ModelReportSource<SimpleModel>(new List<SimpleModel>
            {
                new SimpleModel("Hello", 1),
                new SimpleModel("World", 2)
            });

            Assert.Null(value.GetValue());

            Assert.True(value.NextRow());
            Assert.True(value.NextColumn());

            Assert.Equal("Hello", value.GetValue());

            Assert.True(value.NextColumn());
            Assert.Equal(1d, value.GetValue());

            Assert.False(value.NextColumn());
            Assert.Null(value.GetValue());

            Assert.True(value.NextRow());
            Assert.Null(value.GetValue());

            Assert.True(value.NextColumn());
            Assert.Equal("World", value.GetValue());

            Assert.True(value.NextColumn());
            Assert.Equal(2d, value.GetValue());

            Assert.False(value.NextColumn());
            Assert.False(value.NextRow());

            Assert.Null(value.GetValue());
        }
    }
}