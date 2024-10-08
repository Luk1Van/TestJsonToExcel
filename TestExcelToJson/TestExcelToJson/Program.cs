using System.Reflection;
using TestExcelToJson.DataAccess;
using TestExcelToJson.XLStoJson;
using System.Reflection;
using System;
using Newtonsoft.Json.Linq;
using TestExcelToJson.DModels;
using System.Runtime.Serialization;
using System.Reflection.Metadata;
using System.Data;
using static System.Net.Mime.MediaTypeNames;
using Microsoft.Office.Interop.Excel;

namespace TestExcelToJson
{
    internal class Program
    {
        static void Main(string[] args)
        {
            ExcelAcess excelAcess = new ExcelAcess();
            ReadDataCalculate calcMethods = new ReadDataCalculate();
            //Путь к файлу Excel
            string xlsxPath = "C:\\Users\\ivan.lukich\\Downloads\\Casco_Matrix_Сalculate Тест-кейсы.xlsx";
            //Путь к Json
            string jsonPath = "C:\\Users\\ivan.lukich\\Downloads\\Сalculate.txt";

            excelAcess.xlRange = excelAcess.openExcel(xlsxPath);

            var dataModel = JsonReadWrite.ReadAndDeserialize(jsonPath);

            CalculateModel calcModel = new CalculateModel();
            CalculateModel.Policy policy = new CalculateModel.Policy();
            Type policyType = policy.GetType();


            TestModel testModel = new TestModel();

            TestModel.Policy policyTest = new TestModel.Policy
            {
                driver = new TestModel.Driver()
            };
            Type policyTetsType = policyTest.GetType();
            allPropertiesInNested(testModel, policyTest,  "test", policyTetsType);
            
            

            /*for (int i = 2; i <= excelAcess.xlRange.Rows.Count; i++)
            {
                calcMethods.SetValuesFromXl(dataModel, excelAcess, i);
                string NewJson = JsonConvert.SerializeObject(dataModel, Formatting.Indented);
                File.WriteAllText($"C:\\Users\\ivan.lukich\\Downloads\\Calculate_js\\00{i}.json", NewJson);
            }*/


            excelAcess.closeExcel();
        }

        public static void Test<T>(T dataModel)
        {
            Type type = typeof(T);
            Console.WriteLine("            " + type.FullName);
            foreach (PropertyInfo property in type.GetProperties())
            {
                Console.WriteLine(property.PropertyType);
                if(!property.PropertyType.IsPrimitive && property.PropertyType != typeof(string))
                {
                    //Test(property.GetValue(dataModel));
                    Test(property.PropertyType);
                }
            }

        }

        /*public static void allPropertiesInNested<T1, T2>(T1 baseClass, T2 dataModel)
        {
            Type objType = dataModel.GetType();
            foreach (PropertyInfo property in objType.GetProperties())
            {
                Console.WriteLine(property.ToString());
                if (GetNestedClasses(baseClass).Contains(property.ToString().Split(' ').First()))
                {
                    Console.WriteLine("   " + property.Name);
                    allPropertiesInNested(baseClass, property);
                }
            }
        }*/

        /// <summary>
        /// 
        /// </summary>
        /// <typeparam name="T1"></typeparam>
        /// <typeparam name="T2"></typeparam>
        /// <param name="baseClass">Класс содержащий остальные классы, нужен для нахождения всех внутренних классов</param>
        /// <param name="dataModel">Root класс</param>
        public static void allPropertiesInNested<T1, T2>(T1 baseClass, T2 dataModel,  string testvalue, Type objectType)
        {
            //if(datatype == null) return;
            //Type datatype = dataModel.GetType();
            Dictionary<string, Type> stringType = GetNestedClasses(baseClass);
            Console.WriteLine("                   " + objectType.Name);
            foreach (PropertyInfo property in objectType.GetProperties())
            {
                Console.WriteLine(property.ToString());
                string propertyString = property.ToString().Split(' ').First().Split("[", StringSplitOptions.RemoveEmptyEntries).First();
                if (stringType.ContainsKey(propertyString))
                {
                    var test = property.PropertyType;
                    allPropertiesInNested(baseClass, dataModel, testvalue, property.PropertyType);
                    var type = Nullable.GetUnderlyingType(test);
                }
                else
                {
                    property.SetValue(dataModel, testvalue);
                }
            }
        }

        /// <summary>
        /// Возвращает названия всех классов внутри заданного класса
        /// </summary>
        /// <param name="calcModel"></param>
        /// <returns>string[] - массив названий </returns>
        public static Dictionary<string, Type> GetNestedClasses<T>(T baseClass)
        {
            Dictionary<string, Type> stringClass = new Dictionary<string, Type>();
            Type[] nestedClasses = baseClass.GetType().GetNestedTypes(BindingFlags.Public);
            foreach (Type nestedClass in nestedClasses)
            {
                stringClass.Add(nestedClass.Name.ToString(), nestedClass);
            }
            return stringClass;
        }

        public static bool DelegateToSearchCriteria(MemberInfo objMemberInfo, Object objSearch)
        {
            // Compare the name of the member function with the filter criteria.
            if (objMemberInfo.Name.ToString() == objSearch.ToString())
                return true;
            else
                return false;
        }

    }
}
