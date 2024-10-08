using Newtonsoft.Json;
using TestExcelToJson.DModels;

namespace TestExcelToJson.XLStoJson
{
    public static class JsonReadWrite
    {
        public static CalculateModel.Rootobject ReadAndDeserialize(string filePath)
        {
            return JsonConvert.DeserializeObject<CalculateModel.Rootobject>(File.ReadAllText(filePath));
        }
    }
}
