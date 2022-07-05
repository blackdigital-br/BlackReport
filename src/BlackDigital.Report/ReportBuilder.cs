using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BlackDigital.Report
{
    public abstract class ReportBuilder
    {
        public abstract Task<byte[]> BuildAsync();

        public async Task BuildAsync(Stream stream)
        {
            if (stream == null || !stream.CanWrite)
                throw new ArgumentException("Stream is null or not writable");
            
            var buffer = await BuildAsync();
            await stream.WriteAsync(buffer);
        }

        public async Task BuildAsync(string file)
        {
            await File.WriteAllBytesAsync(file, await BuildAsync());
        }

    }
}
