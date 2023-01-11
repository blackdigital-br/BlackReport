using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BlackDigital.Report
{
    internal struct TemplateKey : IEquatable<TemplateKey>
    {
        public string Name { get; set; }

        public string Type { get; set; }

        public override bool Equals([NotNullWhen(true)] object? obj)
        {
            if (obj is TemplateKey key)
                return key.Equals(key);

            return base.Equals(obj);
        }

        public bool Equals(TemplateKey other)
        {
            return this.Name == other.Name && this.Type == other.Type;
        }

        public override int GetHashCode()
        {
            return HashCode.Combine(Name, Type);
        }
    }
}
