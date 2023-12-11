using BlackDigital.Report.Example.Model;
using BlackDigital.Report.Sources;

namespace BlackDigital.Report.Tests.Sources
{
    public class ReportSourceTest
    {
        [Fact]
        public async void SingleReportValue()
        {
            var value = new SingleReportSource();
            value.Load("Hello World");

            Assert.Null(await value.GetValueAsync());

            Assert.True(await value.NextRowAsync());
            Assert.True(await value.NextColumnAsync());

            Assert.Equal("Hello World", await value.GetValueAsync());

            Assert.False(await value.NextColumnAsync());
            Assert.Null(await value.GetValueAsync());

            Assert.False(await value.NextRowAsync());
            Assert.Null(await value.GetValueAsync());
        }

        [Fact]
        public async void SingleReportValueWithNull()
        {
            var value = new SingleReportSource();

            Assert.Null(await value.GetValueAsync());

            Assert.True(await value.NextRowAsync());
            Assert.True(await value.NextColumnAsync());

            Assert.Null(await value.GetValueAsync());

            Assert.False(await value.NextColumnAsync());
            Assert.Null(await value.GetValueAsync());

            Assert.False(await value.NextRowAsync());
            Assert.Null(await value.GetValueAsync());
        }

        [Fact]
        public async void EnumerableReportValue()
        {
            var value = new EnumerableReportSource();

            value.Load(new[]
            {
                new[] { "Hello", "World" },
                new[] { "Hello2", "World2" }
            });

            Assert.Null(await value.GetValueAsync());

            Assert.True(await value.NextRowAsync());
            Assert.True(await value.NextColumnAsync());

            Assert.Equal("Hello", await value.GetValueAsync());

            Assert.True(await value.NextColumnAsync());
            Assert.Equal("World", await value.GetValueAsync());

            Assert.False(await value.NextColumnAsync());
            Assert.Null(await value.GetValueAsync());

            Assert.True(await value.NextRowAsync());
            Assert.Null(await value.GetValueAsync());

            Assert.True(await value.NextColumnAsync());
            Assert.Equal("Hello2", await value.GetValueAsync());

            Assert.True(await value.NextColumnAsync());
            Assert.Equal("World2", await value.GetValueAsync());

            Assert.False(await value.NextColumnAsync());
            Assert.False(await value.NextRowAsync());

            Assert.Null(await value.GetValueAsync());
        }

        [Fact]
        public async void ModelReportSource()
        {
            var value = new ModelReportSource<SimpleModel>();

            value.Load(new List<SimpleModel>
            {
                new SimpleModel("Hello", 1),
                new SimpleModel("World", 2)
            });

            Assert.Null(await value.GetValueAsync());

            Assert.True(await value.NextRowAsync());
            Assert.True(await value.NextColumnAsync());

            Assert.Equal("Hello", await value.GetValueAsync());

            Assert.True(await value.NextColumnAsync());
            Assert.Equal(1d, await value.GetValueAsync());

            Assert.False(await value.NextColumnAsync());
            Assert.Null(await value.GetValueAsync());

            Assert.True(await value.NextRowAsync());
            Assert.Null(await value.GetValueAsync());

            Assert.True(await value.NextColumnAsync());
            Assert.Equal("World", await value.GetValueAsync());

            Assert.True(await value.NextColumnAsync());
            Assert.Equal(2d, await value.GetValueAsync());

            Assert.False(await value.NextColumnAsync());
            Assert.False(await value.NextRowAsync());

            Assert.Null(await value.GetValueAsync());
        }

        [Fact]
        public async void ListReportSource()
        {
            var list = new List<object>
            {
                "Hello",
                "World"
            };

            var value = new ListReportSource(list);

            Assert.Null(await value.GetValueAsync());

            Assert.True(await value.NextRowAsync());
            Assert.True(await value.NextColumnAsync());

            Assert.Equal("Hello", await value.GetValueAsync());

            Assert.True(await value.NextColumnAsync());
            Assert.Equal("World", await value.GetValueAsync());

            Assert.False(await value.NextColumnAsync());
            Assert.False(await value.NextRowAsync());

            Assert.Null(await value.GetValueAsync());
        }
    }
}