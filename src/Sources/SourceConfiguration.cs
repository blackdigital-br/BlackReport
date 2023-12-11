using System;
using System.Collections.Generic;
using System.ComponentModel.Design;
using System.Linq;

namespace BlackDigital.Report.Sources
{
    public class SourceConfiguration
    {
        internal SourceConfiguration()
        {
            SourcesTypes = new();
            CreateBaseSources();
        }

        private List<Type> SourcesTypes;

        private void CreateBaseSources()
        {
            SourcesTypes.Add(typeof(DataReaderReportSource));
            SourcesTypes.Add(typeof(DataTableReportSource));
            SourcesTypes.Add(typeof(EnumerableReportSource));
            SourcesTypes.Add(typeof(ModelReportSource<>));
            SourcesTypes.Add(typeof(ListReportSource));
            SourcesTypes.Add(typeof(SingleReportSource));
        }

        public SourceConfiguration Add(Type sourceType)
        {
            if (!sourceType.IsSubclassOf(typeof(ReportSource)))
                throw new ArgumentException("Source type must be a subclass of ReportSource");

            SourcesTypes.Add(sourceType);

            return this;
        }

        public SourceConfiguration Add<T>() 
            where T : ReportSource
        {
            SourcesTypes.Add(typeof(T));
            return this;
        }

        public SourceConfiguration Clear()
        {
            SourcesTypes.Clear();
            return this;
        }

        public ReportSource? FindSource(Type type, object? source)
        {
            if (source is ReportSource)
                return source as ReportSource;

            foreach (var sourceType in SourcesTypes)
            {
                object? sourceObject;

                if (sourceType.IsGenericTypeDefinition
                    && type.IsGenericType
                    && type.GenericTypeArguments.Length == 1
                    && type.GenericTypeArguments.First().IsClass
                    && type.GenericTypeArguments.First() != typeof(string))
                    sourceObject = Activator.CreateInstance(sourceType.MakeGenericType(type.GenericTypeArguments));
                else if (sourceType.IsGenericTypeDefinition)
                    sourceObject = Activator.CreateInstance(sourceType.MakeGenericType(type));
                else
                    sourceObject = Activator.CreateInstance(sourceType);

                if (sourceObject is ReportSource reportSource
                    && reportSource.IsSourceType(type, source))
                {
                    reportSource.Load(source);
                    return reportSource;
                }
            }

            return null;
        }

        public ReportSource? FindSource<T>(T? source = default)
            => FindSource(typeof(T), source);
    }
}
