using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Threading.Tasks;

namespace BlackDigital.Report.Sources
{
    public class ModelPaginationReportSource<T> : ReportSource
        where T : class
    {
        #region "Constructors"

        public ModelPaginationReportSource() 
        { 
            Take = 5000;
            Skip = 0;

            Properties = typeof(T).GetProperties();
        }

        public ModelPaginationReportSource(GetDataPaginationFunc getDataPagination)
            : this()
        {
            GetDataPagination = getDataPagination;
        }

        public ModelPaginationReportSource(GetDataPaginationFunc getDataPagination, ProgressAction progress)
            : this(getDataPagination)
        {
            Progress = progress;
        }

        #endregion "Constructors"

        #region "Delegates"

        public delegate Task<IEnumerable<T>> GetDataPaginationFunc(int skip, int take);

        public delegate Task ProgressAction(double progress);

        #endregion "Delegates"

        #region "Properties"

        private uint _rowCount = 0;

        public override uint RowCount => _rowCount;

        private uint _columnCount = 0;
        public override uint ColumnCount => _columnCount;

        public int Take { get; set; }

        public int Skip { get; set; }

        public int? TotalItens { get; set; }

        private List<List<object>>? _data = null;

        private int _rowPosition = -1;
        private int _columnPosition = -1;
        private double _progress = 0.0;


        protected IEnumerable<PropertyInfo>? Properties;

        public GetDataPaginationFunc? GetDataPagination { get; set; }

        public ProgressAction? Progress { get; set; }

        #endregion "Properties"

        #region "ReportSource"

        public override bool IsSourceType(Type type, object? value)
        {
            return value is GetDataPaginationFunc;
        }

        public override void Load(object data) 
        { 
            if (data == null)
                throw new NullReferenceException("Data cannot be null");

            if (!IsSourceType(data.GetType(), data))
                throw new Exception("Invalid data type");

            GetDataPagination = (GetDataPaginationFunc)data;
        }
        
        private async Task<bool> LoadDataAsync()
        {
            if (GetDataPagination == null)
                throw new NullReferenceException("GetDataPagination cannot be null");

            if (_data != null)
            {
                if (_data.Count < Take)
                    return false;

                GC.SuppressFinalize(_data);
                _data = null;
                GC.Collect();
            }
                

            var data = await GetDataPagination(Skip, Take);
            Skip += Take;

            if (data == null || !data.Any())
            {
                _data = null;
                return false;
            }

            _data = new List<List<object>>();
            var properties = ReportHelper.GetPropertiesAndAttributes<T>();

            foreach (var row in data)
            {
                var dataRow = new List<object>();

                foreach (var property in properties)
                {
                    dataRow.Add(property.Item1?.GetValue(row) ?? string.Empty);
                }

                _data.Add(dataRow);
            }

            _rowPosition = 0;

            return true;
        }

        public override async Task<bool> NextRowAsync()
        {
            _rowPosition++;

            if (_data == null || _rowPosition >= _data.Count())
            {
                var hasData = await LoadDataAsync();

                if (!hasData)
                    return false;
            }

            if (Progress != null
                    && TotalItens != null
                    && TotalItens > 0)
            {
                double percent = (double)_rowCount / (double)TotalItens.Value * 100;
                percent = Math.Round(percent, 2);

                if ((percent - _progress) > 1)
                {
                    _progress = percent;
                    await Progress(percent);
                }
            }

            _rowCount++;
            _columnPosition = -1;

            return true;
        }

        public override Task<bool> NextColumnAsync()
        {
            if (_data == null || _rowPosition < 0)
                throw new Exception("ReportSource is not loaded");

            _columnPosition++;

            if (_columnPosition >= (Properties?.Count() ?? 0))
                return Task.FromResult(false);

            int columnSize = _columnPosition + 1;

            if (columnSize > _columnCount)
                _columnCount = (uint)columnSize;

            return Task.FromResult(true);
        }

        public override Task<object?> GetValueAsync()
        {
            if (_data == null)
                throw new Exception("ReportSource is not loaded");

            if (_rowPosition < 0 || _columnPosition < 0)
                throw new Exception("ReportSource is not loaded");

            if (_rowPosition >= _data.Count())
                throw new Exception("Row position is out of range");

            if (_columnPosition >= (Properties?.Count() ?? 0))
                throw new Exception("Column position is out of range");

            var row = _data.ElementAt(_rowPosition);
            var value = row?.ElementAt(_columnPosition) ?? null;

            return Task.FromResult(value);
        }

        public override Task ResetAsync()
        {
            _data = null;
            _rowPosition = -1;
            _columnPosition = -1;
            _rowCount = 0;
            _columnCount = 0;
            Skip = 0;

            return Task.CompletedTask;
        }

        #endregion "ReportSource"
    }
}
