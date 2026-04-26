using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TestTask3
{
    internal interface IWriter
    {
        IDictionary<string, string> GetParams();
        void SetParams(IDictionary<string, string> parameters);
    }
}
