using System;

namespace NpoiMapper
{
    public class ImportMatchException : Exception
    {
        public int RowIndex { get; private set; }
        public int ColumnIndex { get; private set; }
        public string FileName { get; private set; }

        public ImportMatchException(int rowIndex, int columnIndex, string fileName, string message)
            : this(rowIndex, columnIndex, fileName, message, null)
        {
        }

        public ImportMatchException(int rowIndex, int columnIndex, string fileName, Exception innerException)
            : this(rowIndex, columnIndex, fileName, null, innerException)
        {
        }

        public ImportMatchException(int rowIndex, int columnIndex, string fileName, string message = null,
            Exception innerException = null)
            : base(message, innerException)
        {
            RowIndex = rowIndex;
            ColumnIndex = columnIndex;
            FileName = fileName;
        }
    }
}