using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TestTask3
{
    internal interface ITableLoader
    {
        DataTable LoadTable(string tableName);
        bool SaveTable(DataTable dataTable, string tableName);
    }
}
