using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using OfficeOpenXml;
using UniLoader.Config;

namespace UniLoader.DataParsers
{
    /// <summary>
    /// Парсер таблиц формата Excel (.xlsx)
    /// </summary>
    public class XlsxDataParser : DataParser
    {
        public XlsxDataParser(string FileName, Table workTable) : base(FileName, workTable)
        {
        }

        public override DataTable ReadFile()
        {
            DataTable dataTable = new DataTable();

            using (var fileStream = new FileStream(fileName, FileMode.Open))
            {
                using (var package = new ExcelPackage(fileStream))
                {
                    var workSheet = package.Workbook.Worksheets[ConfTable.Worksheet];
                    if (workSheet == null) return dataTable;
                    var columns = GetHeader(workSheet);

                    foreach (var column in columns) dataTable.Columns.Add(column);

                    for (int rowIndex = ConfTable.StartRow; rowIndex <= workSheet.Dimension.End.Row; rowIndex++)
                    {
                        var tempRow = dataTable.NewRow();

                        for (int columnIndex = ConfTable.StartColumn; columnIndex <= workSheet.Dimension.End.Column; columnIndex++)
                        {
                            var value = workSheet.Cells[rowIndex, columnIndex].Text;  
                            var dtColIndex = columnIndex - ConfTable.StartColumn;
                            tempRow[dtColIndex] = ParseValue(value, tempRow.Table.Columns[dtColIndex]);
                        }
                        dataTable.Rows.Add(tempRow);
                    }
                }
            }

            IEnumerable<DataColumn> queryableColumns = dataTable.Columns.Cast<DataColumn>().AsQueryable();
            var columnsToRemove = queryableColumns.Where(column => column.ColumnName != ConfTable.Columns
                                     .FirstOrDefault(x => x.ExcelName == column.ColumnName).ExcelName);
            foreach (var column in columnsToRemove.ToList())
            {
                dataTable.Columns.Remove(column);
                Logger.WriteLine($"Столбец '{column.ColumnName}' был удален из набора данных, поскольку его нет в конфиге.");
            }

            if(dataTable.Rows.Count == 0) { Logger.WriteLine("Файл не содержит записей.");}
            dataTable = dataTable.Rows.Cast<DataRow>().Where(row =>
                        !row.ItemArray.All(field => field is DBNull || string.IsNullOrWhiteSpace(field as string)))
                    .CopyToDataTable();

                Console.Write($"({dataTable.Rows.Count} записей) : ");
            return dataTable;
        }

        /// <summary>
        /// Распарсить значение согласно типу данных столбца
        /// </summary>
        private object ParseValue(string value, DataColumn column)
        {
            if (value == string.Empty)  
            {
                return DBNull.Value;
            }
            else
            {
                if (column.DataType == typeof(DateTime)) //проверка дат
                {
                    DateTime date;
                    if (DateTime.TryParse(value, out date))
                    {
                        return value;
                    }
                    else
                    {
                        return (new DateTime(1899, 12, 30).AddDays(Convert.ToInt32(value))).ToShortDateString();
                    }
                }
                else if (column.DataType == typeof(int)) //проверка целых чисел
                {
                    int num;
                    if (int.TryParse(value, out num))
                    {
                        return num;
                    }
                    else
                    {
                        throw new Exception("TypeError in file");
                    }
                }
                else if (column.DataType == typeof(double)) //проверка даблов
                {
                    double num;
                    if (double.TryParse(value, out num))
                    {
                        return num;
                    }
                    else
                    {
                        throw new Exception("TypeError in file");
                    }
                }
                else
                {
                    value = TrimVarchar(value, column);
                    return value;
                }
            }
        }

        private List<DataColumn> GetHeader(ExcelWorksheet workSheet)
        {
            var cells = workSheet.Cells[ConfTable.HeaderRow, ConfTable.StartColumn, 1, workSheet.Dimension.End.Column];
            List<DataColumn> header = new List<DataColumn>();

            foreach (var cell in cells)
            {
                header.Add(new DataColumn(cell.Text));
            }

            foreach (var column in header)
            {
                var item = ConfTable.Columns.FirstOrDefault(x => x.ExcelName == column.ColumnName);
                if (item.DbType!= null) column.DataType = GetColumnType(item.DbType);
            }
            return header;
        }

        protected override Type GetColumnType(string DbType)
        {
            DbType = DbType.ToUpper();
            if (DbType.Contains("VARCHAR"))
            {
                return typeof(string);
            }
            else if (DbType.Contains("NUMBER"))
            {
                return typeof(double);
            }
            else if (DbType.Contains("INTEGER"))
            {
                return typeof(int);
            }
            else if (DbType.Contains("DATE"))
            {
                return typeof(DateTime);
            }
            return typeof(string);
        }

        private string TrimVarchar(string value, DataColumn column)
        {
            var confColumn = ConfTable.Columns.FirstOrDefault(x => x.ExcelName == column.ColumnName);
            if (confColumn.DbType == null) return value;

            var dbType = confColumn.DbType.ToUpper();
            if (dbType.Contains("VARCHAR"))
            {
                var size = Convert.ToInt16(dbType.Split('(', ')')[1]);
                if (value.Length > size)
                {
                    Logger.WriteLine($"ВНИМАНИЕ: Обнаружено значение текстового поля ({column.ColumnName} ({value.Length} символов)), длина которого превышает допустимое ({size} символов).");
                    Logger.WriteLine("Его длина была сокращена до требуемой.");

                    value = value.Substring(0, size);
                }
            }
            return value;
        }
    }
}
