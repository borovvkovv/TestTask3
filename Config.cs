using System;
using System.Collections.Generic;
using System.Dynamic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TestTask3
{
    internal class Config(IWriter writer)
    {
        private readonly IDictionary<string, string> _params = writer.GetParams();

        public bool TryGetParam(string paramName, out string paramValue)
        {
            paramValue = _params.Where(p => p.Key == paramName).Select(p => p.Value).FirstOrDefault();

            return _params.Any(p => p.Key == paramName);
        }

        public void SetConfigParam(string paramName, string paramValue)
        {
            if (_params.ContainsKey(paramName))
                _params[paramName] = paramValue;
            else
                _params.Add(paramName, paramValue);
            
            writer.SetParams(_params);
        }
    }
}
