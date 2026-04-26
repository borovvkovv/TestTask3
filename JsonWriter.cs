using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TestTask3
{
    internal class JsonWriter(string filePath): IWriter
    {
        private readonly string _fullFilePath = Path.Combine(AppContext.BaseDirectory, filePath);

        public IDictionary<string, string> GetParams()
        {
            var result = new Dictionary<string, string>();
            var jsonText = string.Empty;
            try
            {
                jsonText = File.ReadAllText(_fullFilePath);
            }
            catch (Exception e)
            {
                return result;
            }
            
            var jsonObj = JsonConvert.DeserializeObject(jsonText);
            if (jsonObj is JObject)
                return JsonConvert.DeserializeObject<Dictionary<string, string>>(jsonText);

            return result;
        }

        public void SetParams(IDictionary<string, string> parameters)
        {
            try
            {
                File.WriteAllText(_fullFilePath, JsonConvert.SerializeObject(parameters));
            }
            catch (Exception e)
            {
                return;
            }
        }
    }
}
