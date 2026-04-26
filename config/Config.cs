using TestTask3.interfaces;

namespace TestTask3.config
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
