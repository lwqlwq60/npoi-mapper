using NPOI.SS.UserModel;

namespace NpoiMapper
{
    public static class MapExtensions
    {
        public static ExportMapper<T> FontSize<T>(this ExportMapper<T> mapper,
            int columnIndex, double fontSize, bool header = true) where T : class
        {
            var column = mapper.CreateOrAddColumnInfo(columnIndex);
            if (header)
                column.HeaderFontSize = fontSize;
            else
                column.DataFontSize = fontSize;
            return mapper;
        }

        public static ExportMapper<T> FontSize<T>(this ExportMapper<T> mapper, double fontSize,
            bool header = true, params int[] columnIndexes) where T : class
        {
            foreach (var columnIndex in columnIndexes)
            {
                mapper.FontSize(columnIndex, fontSize, header);
            }
            return mapper;
        }

        public static ExportMapper<T> Format<T>(this ExportMapper<T> mapper, int columnIndex,
            string format) where T : class
        {
            var column = mapper.CreateOrAddColumnInfo(columnIndex);
            column.DataFormat = format;
            return mapper;
        }

        public static ExportMapper<T> Format<T>(this ExportMapper<T> mapper, string format
            , params int[] columnIndexes) where T : class
        {
            foreach (var columnIndex in columnIndexes)
            {
                mapper.Format(columnIndex, format);
            }
            return mapper;
        }

        public static ExportMapper<T> Bold<T>(this ExportMapper<T> mapper, int columnIndex,
            bool bold = true, bool header = true) where T : class
        {
            var column = mapper.CreateOrAddColumnInfo(columnIndex);
            if (header)
                column.HeaderIsBold = bold;
            else
                column.DataIsBold = bold;
            return mapper;
        }

        public static ExportMapper<T> Bold<T>(this ExportMapper<T> mapper, bool bold = true,
            bool header = true, params int[] columnIndexes) where T : class
        {
            foreach (var columnIndex in columnIndexes)
            {
                mapper.Bold(columnIndex, bold, header);
            }
            return mapper;
        }

        public static ExportMapper<T> HorizontalAlignment<T>(this ExportMapper<T> mapper,
            int columnIndex, HorizontalAlignment alignment, bool header = true) where T : class
        {
            var column = mapper.CreateOrAddColumnInfo(columnIndex);
            if (header)
                column.HeaderHorizontalAlignment = alignment;
            else
                column.DataHorizontalAlignment = alignment;
            return mapper;
        }

        public static ExportMapper<T> HorizontalAlignment<T>(this ExportMapper<T> mapper,
            HorizontalAlignment alignment, bool header = true, params int[] columnIndexes) where T : class
        {
            foreach (var columnIndex in columnIndexes)
            {
                mapper.HorizontalAlignment(columnIndex, alignment, header);
            }
            return mapper;
        }

        public static ExportMapper<T> VerticalAlignment<T>(this ExportMapper<T> mapper,
            int columnIndex, VerticalAlignment alignment, bool header = true) where T : class
        {
            var column = mapper.CreateOrAddColumnInfo(columnIndex);
            if (header)
                column.HeaderVerticalAlignment = alignment;
            else
                column.DataVerticalAlignment = alignment;
            return mapper;
        }

        public static ExportMapper<T> VerticalAlignment<T>(this ExportMapper<T> mapper,
            VerticalAlignment alignment, bool header = true, params int[] columnIndexes) where T : class
        {
            foreach (var columnIndex in columnIndexes)
            {
                mapper.VerticalAlignment(columnIndex, alignment, header);
            }
            return mapper;
        }
    }
}
