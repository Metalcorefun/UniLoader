using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using Oracle.ManagedDataAccess.Client;

namespace UniLoader.DataClients
{
    class OracleDataClient : DataClient
    {
        private OracleConnection connection;
        public OracleDataClient(string connectionString) : base(connectionString)
        {
            connection = new OracleConnection(ConnectionString);
        }

        public override bool TestConnection()
        {
            try { connection.Open(); }
                catch (Exception) { return false; }
                connection.Close();
                return true;
        }

        public override void Send(DataTable dataTable) //Делаем bulk-insert
        {
            #region Vars
            var queryString = BuildInsertQuery();

            var oracleCommand = new OracleCommand(queryString, connection);
            var oracleParameters = new List<OracleParameter>();

            var data = ExtractData(dataTable);
            var dates = CreateCurrentDateArray(dataTable.Rows.Count);
            #endregion

            
            for (int i = 0; i < data.Count; i++)
            {
                var parameter = new OracleParameter();
                parameter.Value = data[i];
                oracleParameters.Add(parameter);
            }

            if (ConfTable.InsertCurrentDate)
            {
                var parameter = new OracleParameter();
                parameter.Value = dates;
                oracleParameters.Add(parameter);
            }

            oracleCommand.ArrayBindCount = data.First().Length;
            foreach (var parameter in oracleParameters)
            {
                oracleCommand.Parameters.Add(parameter);
            }

            connection.Open();
            try
            {
                oracleCommand.ExecuteNonQuery();
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
                throw;
            }
            finally { connection.Close();}

            string message = $"Загружено {data.First().Length} записей.\n";
            Console.WriteLine(" Ок!");
            Logger.WriteLine(message);
        }

        private string BuildInsertQuery()
        {
            var queryString = $"INSERT INTO {ConfTable.SchemaInUse}.{ConfTable.TableName}";
            string param = "", values = "";

            for (int columnIndex = 0; columnIndex < ConfTable.Columns.Count; columnIndex++)
            {
                param += $"{ConfTable.Columns.ElementAt(columnIndex).DbName}, ";
                values += $":{columnIndex + 1}, ";
            }

            if (ConfTable.InsertCurrentDate) { param += "INSERT_DATE, "; values += $":{ConfTable.Columns.Count + 1}"; }

            param = param.EndsWith(", ") ? param.Substring(0, param.Length - 2) : param;
            values = values.EndsWith(", ") ? values.Substring(0, values.Length - 2) : values;

            queryString += $" ({param}) VALUES ({values})";
            return queryString;
        }

        private List<string[]> ExtractData(DataTable table)
        {
            List<string[]> data = (from object column in table.Columns select new string[table.Rows.Count]).ToList();

            for (int rowIndex = 0; rowIndex < table.Rows.Count; rowIndex++)
            {
                foreach (DataColumn column in table.Columns)
                {
                    var columnName = column.ColumnName;
                    var columnIndex = ConfTable.Columns.FindIndex(x => x.ExcelName == columnName);

                    DateTime dateData;
                    if (DateTime.TryParse(Convert.ToString(table.Rows[rowIndex][table.Columns[columnName].Ordinal]),
                        out dateData))
                    {
                        data[columnIndex][rowIndex] = dateData.ToString("dd.MM.yyy");
                    }
                    else
                    {
                        data[columnIndex][rowIndex] =
                            Convert.ToString(table.Rows[rowIndex][table.Columns[columnName].Ordinal]);
                    }
                }
            }
            return data;
        }

        private string[] CreateCurrentDateArray(int length)
        {
            var dates = new string[length];

            for (int i = 0; i < length; i++) 
            {
                dates[i] = DateTime.Today.ToString("dd.MM.yy");
            }

            return dates;
        }
    }
}
