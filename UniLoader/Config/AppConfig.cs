using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;

namespace UniLoader.Config
{
    /// <summary>
    /// Класс-обёртка вокруг кофигурационного файла
    /// </summary>
    public static class AppConfig
    {
        public static string DatabaseConnectionString { get; set; }
        //Прилетела птица и забрала строку подключения OLEDB в лучший мир...
        public static string WorkingDirectory { get; set; }
        public static string FileExtension { get; set; }
        public static bool DeleteFilesAfterLoad { get; set; }

        public static ICollection<Table> Tables { get; set; }

        public static void Check()
        {
            Exception e = new ArgumentNullException("Config file is corrupted.");
            if (string.IsNullOrWhiteSpace(DatabaseConnectionString)) throw e;
            if (string.IsNullOrWhiteSpace(WorkingDirectory)) throw e;
            if (string.IsNullOrWhiteSpace(FileExtension)) throw e;
        }
    }

    public struct Table
    {
        
        public string TableName { get; set; }
        public string Worksheet { get; set; }
        public string SchemaInUse { get; set; }
        public string SubDirectory { get; set; }
        public int HeaderRow { get; set; }
        public int StartRow { get; set; }
        public int StartColumn { get; set; }
        public bool InsertCurrentDate { get; set; }
        public List<Column> Columns { get; set; }

        public override bool Equals(object obj)
        {
            if (!(obj is Table))
                return false;

            Table table = (Table)obj;

            var matching = this.Columns.Intersect(table.Columns).ToList();

            return this.TableName == table.TableName
                   && this.SubDirectory == table.SubDirectory
                   && this.SchemaInUse == table.SchemaInUse
                   && this.Worksheet == table.Worksheet
                   && this.HeaderRow == table.HeaderRow
                   && this.StartColumn == table.StartColumn
                   && this.StartRow == table.StartRow
                   && this.InsertCurrentDate == table.InsertCurrentDate
                   && this.Columns.Count == table.Columns.Count
                   && this.Columns.Count == matching.Count;
        }
    }

    public struct Column
    {
        public string ExcelName { get; set; }
        public string DbName { get; set; }
        public string DbType { get; set; }
        public Column(string excelName, string dbName, string dbType)
        {
            ExcelName = excelName;
            DbName = dbName;
            DbType = dbType;
        }
    }
}
