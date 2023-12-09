using BlackDigital.Report.Example.Model;
using BlackDigital.Report.Sources;

namespace BlackDigital.Report.Tests.Sources
{
    public class ReportSourceTest
    {
        [Fact]
        public void SingleReportValue()
        {
            var value = new SingleReportSource();
            value.Load("Hello World");

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
            var value = new SingleReportSource();

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
            var value = new EnumerableReportSource();

            value.Load(new[]
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
            var value = new ModelReportSource<SimpleModel>();

            value.Load(new List<SimpleModel>
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

        [Fact]
        public void ListReportSource()
        {
            var list = new List<object>
            {
                "Hello",
                "World"
            };

            var value = new ListReportSource(list);

            Assert.Null(value.GetValue());

            Assert.True(value.NextRow());
            Assert.True(value.NextColumn());

            Assert.Equal("Hello", value.GetValue());

            Assert.True(value.NextColumn());
            Assert.Equal("World", value.GetValue());

            Assert.False(value.NextColumn());
            Assert.False(value.NextRow());

            Assert.Null(value.GetValue());
        }
    }
}