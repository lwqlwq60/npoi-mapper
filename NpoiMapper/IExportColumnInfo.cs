using System;
using NPOI.SS.UserModel;

namespace NpoiMapper
{
    /// <summary>
    /// Information required for one column when mapping between object and file rows.
    /// </summary>
    public interface IExportColumnInfo<T>
    {
        object HeaderValue { get; set; }

        string DataFormat { get; set; }

        bool HeaderIsBold { get; set; }

        bool DataIsBold { get; set; }

        double HeaderFontSize { get; set; }

        double DataFontSize { get; set; }

        HorizontalAlignment? HeaderHorizontalAlignment { get; set; }

        VerticalAlignment? HeaderVerticalAlignment { get; set; }

        HorizontalAlignment? DataHorizontalAlignment { get; set; }

        VerticalAlignment? DataVerticalAlignment { get; set; }

        Func<T, object> GetDataValue { get; set; }

    }
}