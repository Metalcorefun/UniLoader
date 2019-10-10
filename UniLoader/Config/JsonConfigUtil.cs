using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Newtonsoft.Json.Linq;
using UniLoader.Config;

namespace UniLoader.Config
{
    public class JsonConfigUtil : ConfigUtil
    {
        private readonly string _json;

        public JsonConfigUtil(string fileName) : base(fileName)
        {
            using (FileStream fileStream = File.OpenRead(FileName))
            {
                byte[] byteStream = new byte[fileStream.Length];
                fileStream.Read(byteStream, 0, byteStream.Length);
                _json = Encoding.Default.GetString(byteStream);
            }
        }

        /// <summary>
        /// Чтение данных из JSON-конфига
        /// </summary>
        public override void ReadConfig()
        {
            try
            {
                var jsonObject = JObject.Parse(_json);
                AppConfig.DatabaseConnectionString = (string) jsonObject.SelectToken("DatabaseConnectionString");
                AppConfig.WorkingDirectory = (string) jsonObject.SelectToken("WorkingDirectory");
                AppConfig.FileExtension = (string) jsonObject.SelectToken("FileExtension");
                AppConfig.DeleteFilesAfterLoad = (bool) jsonObject.SelectToken("DeleteFilesAfterLoad");
                JArray jsonTables = (JArray) jsonObject["Tables"];

                //Формирование таблиц
                ICollection<Table> tables = new List<Table>();
                foreach (var jsonTable in jsonTables)
                {
                    Table tempTable = new Table();
                    tempTable.TableName = (string) jsonTable.SelectToken("TableName");
                    tempTable.Worksheet = (string) jsonTable.SelectToken("Worksheet");
                    tempTable.SchemaInUse = (string) jsonTable.SelectToken("SchemaInUse");
                    tempTable.SubDirectory = (string) jsonTable.SelectToken("SubDirectory");
                    tempTable.HeaderRow = (int) jsonTable.SelectToken("HeaderRow");
                    tempTable.StartRow = (int) jsonTable.SelectToken("StartRow");
                    tempTable.StartColumn = (int) jsonTable.SelectToken("StartColumn");
                    tempTable.InsertCurrentDate = (bool) jsonTable.SelectToken("InsertCurrentDate");

                    List<Column> columns = new List<Column>();
                    JArray jsonColumns = (JArray) jsonTable["Columns"];
                    foreach (var jsonColumn in jsonColumns)
                    {
                        Column tempColumn = new Column();
                        tempColumn.ExcelName = (string) jsonColumn.SelectToken("ExcelName");
                        tempColumn.DbName = (string) jsonColumn.SelectToken("DBName");
                        tempColumn.DbType = (string) jsonColumn.SelectToken("DBType");
                        columns.Add(tempColumn);
                    }

                    tempTable.Columns = columns;
                    tables.Add(tempTable);
                }
                AppConfig.Tables = tables;
            }
            catch (Exception)
            {
                throw new ArgumentNullException("Config file is corrupted.");
            }
            AppConfig.Check();
        }
    }
}
