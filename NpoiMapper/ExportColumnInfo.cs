using System;
using NPOI.SS.UserModel;

namespace NpoiMapper
{
    public class ExportColumnInfo<T> : IExportColumnInfo<T>
    {
        public ExportColumnInfo()
        {
            SetDefaultStyle();
        }
        private void SetDefaultStyle()
        {
            HeaderIsBold = true;
            DataIsBold = false;
            HeaderFontSize = 12;
            DataFontSize = 11;
            HeaderHorizontalAlignment = HorizontalAlignment.Center;
        }
        public object HeaderValue { get; set; }
        public string DataFormat { get; set; }
        public bool HeaderIsBold { get; set; }
        public bool DataIsBold { get; set; }
        public double HeaderFontSize { get; set; }
        public double DataFontSize { get; set; }
        public HorizontalAlignment? HeaderHorizontalAlignment { get; set; }
        public VerticalAlignment? HeaderVerticalAlignment { get; set; }
        public HorizontalAlignment? DataHorizontalAlignment { get; set; }
        public VerticalAlignment? DataVerticalAlignment { get; set; }
        public Func<T, object> GetDataValue { get; set; }
    }
}
