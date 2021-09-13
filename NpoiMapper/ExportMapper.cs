using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace NpoiMapper
{
    /// <summary>
    /// export object-excel mapper
    /// </summary>
    /// <typeparam name="T">model</typeparam>
    public sealed class ExportMapper<T>
    {
        private readonly IDictionary<int, IExportColumnInfo<T>> _mapCache;

        private IWorkbook _workbook;

        public ExportMapper()
        {
            _mapCache = new Dictionary<int, IExportColumnInfo<T>>();
        }

        public ExportMapper<T> Map(int columnIndex, Func<T, object> getter, object header)
        {
            var column = CreateOrAddColumnInfo(columnIndex);
            column.HeaderValue = header;
            column.GetDataValue = getter;
            return this;
        }

        public void Save(Stream stream, IEnumerable<T> objects, string sheetName = "sheet1", bool xlsx = true)
        {
            _workbook = xlsx ? new XSSFWorkbook() : (IWorkbook) new HSSFWorkbook();
            var sheet = _workbook.GetSheet(sheetName) ?? _workbook.CreateSheet(sheetName);
            CreateHeader(sheet);
            var objectArray = objects.AsMapSource<T>() as T[] ?? objects.ToArray();
            var rowIndex = sheet.FirstRowNum + 1;
            var styleCache = new Dictionary<int, ICellStyle>();
            var fontCache = new Dictionary<int, IFont>();
            foreach (var o in objectArray)
            {
                var row = sheet.CreateRow(rowIndex);
                foreach (var key in _mapCache.Keys)
                {
                    ICellStyle style = null;
                    IFont font = null;
                    CreateOrGetFontAndStyle(key, fontCache, styleCache, ref font, ref style);
                    var column = _mapCache[key];
                    var cell = row.CreateCell(key);

                    if (!string.IsNullOrWhiteSpace(column.DataFormat))
                        style.DataFormat = _workbook.CreateDataFormat().GetFormat(column.DataFormat);
                    font.FontHeightInPoints = column.DataFontSize;
                    if (column.DataHorizontalAlignment.HasValue)
                        style.Alignment = column.DataHorizontalAlignment.Value;
                    if (column.DataVerticalAlignment.HasValue)
                        style.VerticalAlignment = column.DataVerticalAlignment.Value;

                    font.IsBold = column.DataIsBold;
                    cell.CellStyle = style;
                    cell.CellStyle.SetFont(font);
                    SetCellValue(cell, column.GetDataValue(o));
                }

                rowIndex++;
            }

            foreach (var key in _mapCache.Keys) sheet.AutoSizeColumn(key);
            _workbook.Write(stream);
        }

        public void Save(string filePath, IEnumerable<T> objects, string sheetName = "sheet1")
        {
            using (var fs = new FileStream(filePath, FileMode.Create, FileAccess.Write))
            {
                var xlsx = filePath.EndsWith(".xlsx");
                Save(fs, objects, sheetName, xlsx);
            }
        }

        public void Clear()
        {
            if (_mapCache != null)
                _mapCache.Clear();
        }

        internal IExportColumnInfo<T> CreateOrAddColumnInfo(int columnIndex)
        {
            if (_mapCache.ContainsKey(columnIndex))
            {
                return _mapCache[columnIndex];
            }

            var column = new ExportColumnInfo<T>();
            _mapCache.Add(columnIndex, column);
            return column;
        }

        private void CreateOrGetFontAndStyle(int columnIndex, IDictionary<int, IFont> fontCache,
            IDictionary<int, ICellStyle> styleCache, ref IFont font, ref ICellStyle style)
        {
            if (!fontCache.ContainsKey(columnIndex))
            {
                font = _workbook.CreateFont();
                font.FontHeightInPoints = 11;
                fontCache.Add(columnIndex, font);
            }
            else
                font = fontCache[columnIndex];

            if (!styleCache.ContainsKey(columnIndex))
            {
                style = _workbook.CreateCellStyle();
                styleCache.Add(columnIndex, style);
            }
            else
                style = styleCache[columnIndex];
        }

        private void CreateHeader(ISheet sheet)
        {
            var row = sheet.CreateRow(sheet.FirstRowNum);

            foreach (var key in _mapCache.Keys)
            {
                IFont font = _workbook.CreateFont();
                font.FontHeightInPoints = 11;

                ICellStyle style = _workbook.CreateCellStyle();

                var cell = row.CreateCell(key);
                var column = _mapCache[key];
                font.FontHeightInPoints = column.HeaderFontSize;
                if (column.HeaderHorizontalAlignment.HasValue)
                    style.Alignment = column.HeaderHorizontalAlignment.Value;
                if (column.HeaderVerticalAlignment.HasValue)
                    style.VerticalAlignment = column.HeaderVerticalAlignment.Value;

                font.IsBold = column.HeaderIsBold;
                cell.CellStyle = style;
                cell.CellStyle.SetFont(font);
                SetCellValue(cell, column.HeaderValue);
            }
        }

        private void SetCellValue(ICell cell, object obj)
        {
            if (obj == null)
            {
                return;
            }
            else if (obj.GetType().IsNumeric())
            {
                cell.SetCellValue(Convert.ToDouble(obj));
            }
            else if (obj is string)
            {
                cell.SetCellValue(obj.ToString());
            }
            else if (obj is DateTime)
            {
                cell.SetCellValue((DateTime) obj);
            }
            else if (obj is bool)
            {
                cell.SetCellValue((bool) obj);
            }
            else if (obj.GetType() == typeof(IRichTextString))
            {
                cell.SetCellValue((IRichTextString) obj);
            }
            else
            {
                cell.SetCellValue(obj.ToString());
            }
        }
    }
}