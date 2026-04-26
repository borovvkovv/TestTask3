using System.Data;

namespace TestTask3.interfaces
{
    internal interface ITableLoader
    {
        DataTable LoadTable(string tableName);
        bool SaveTable(DataTable dataTable, string tableName);
    }
}
