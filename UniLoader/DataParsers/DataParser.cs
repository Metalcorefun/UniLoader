using System;
using System.Data;
using UniLoader.Config;

namespace UniLoader.DataParsers
{
    /// <summary>
    /// Базовый класс парсера таблиц.
    /// </summary>
    public abstract class DataParser
    {
        public string fileName { get; set; }
        public Table ConfTable { get; set; }

        protected DataParser(string FileName, Table workTable)
        {
            fileName = FileName;
            ConfTable = workTable;
        }

        public abstract DataTable ReadFile();

        protected abstract Type GetColumnType(string DbType);
    }
}
