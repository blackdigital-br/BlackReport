using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Resources;
using System.Globalization;
using System.Threading.Tasks;

namespace BlackDigital.Report
{
    public abstract class ReportBuilder
    {
        #region "Properties"

        internal ResourceManager? Resource { get; private set; }

        internal IFormatProvider? FormatProvider { get; private set; }

        #endregion "Properties"

        #region "Builder"

        protected TBuilder SetResourceManager<TBuilder>(ResourceManager resource)
            where TBuilder : ReportBuilder
        {
            Resource = resource;
            return (TBuilder)this;
        }

        protected TBuilder SetFormatProvider<TBuilder>(IFormatProvider formatProvider)
            where TBuilder : ReportBuilder
        {
            FormatProvider = formatProvider;
            return (TBuilder)this;
        }

        #endregion "Builder"

        #region "Build"

        public abstract Task<byte[]> BuildAsync();

        public virtual async Task BuildAsync(Stream stream)
        {
            if (stream == null || !stream.CanWrite)
                throw new ArgumentException("Stream is null or not writable");
            
            var buffer = await BuildAsync();
            await stream.WriteAsync(buffer);
        }

        public virtual async Task BuildAsync(string file)
        {
            await File.WriteAllBytesAsync(file, await BuildAsync());
        }

        public virtual void SaveAsTemplate(string name)
        {
            BuilderTemplate.AddTemplate(name, this);
        }

        #endregion "Build"
    }
}
