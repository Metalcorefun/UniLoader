using System.Data;
using UniLoader.Config;

namespace UniLoader.DataClients
{
    public abstract class DataClient
    {
        protected string ConnectionString { get; set; }
        public Table ConfTable { get; set; }

        public DataClient(string connectionString) => ConnectionString = connectionString;

        public abstract bool TestConnection();
        public abstract void Send(DataTable dataTable);
    }
}
