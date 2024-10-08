using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using TestExcelToJson.DataAccess;
using TestExcelToJson.DModels;

namespace TestExcelToJson.XLStoJson
{
    public class ReadDataCalculate
    {
        public string CheckIfNull(ExcelAcess exAcces, int row, int column)
        {
            if (exAcces.xlRange.Cells[row, column].Value2 != null)
            {
                return exAcces.xlRange.Cells[row, column].Value2.ToString();
            }
            return "";
        }

        public void SetValuesFromXl(CalculateModel.Rootobject dataModel, ExcelAcess excelAcess, int i)
        {
            dataModel.policy.AmountCurrency = excelAcess.xlRange.Cells[i, 2].Value2;
            dataModel.policy.AmountCurrencyName = excelAcess.xlRange.Cells[i, 3].Value2;
            dataModel.policy.BSOType = excelAcess.xlRange.Cells[i, 4].Value2;
            dataModel.policy.DocumentDate = CheckIfNull(excelAcess, i, 5);
            dataModel.policy.Drivers[0].Age = Convert.ToInt32(excelAcess.xlRange.Cells[i, 7].Value2);
            dataModel.policy.Drivers[0].Birthdate = excelAcess.xlRange.Cells[i, 8].Value2;
            dataModel.policy.Drivers[0].DriverInsurer = StringToBool(excelAcess.xlRange.Cells[i, 9].Value2);
            dataModel.policy.Drivers[0].DriverOwner = StringToBool(excelAcess.xlRange.Cells[i, 10].Value2);
            dataModel.policy.Drivers[0].Experience = Convert.ToInt32(excelAcess.xlRange.Cells[i, 11].Value2);
            dataModel.policy.Drivers[0].Gender = excelAcess.xlRange.Cells[i, 12].Value2;
            dataModel.policy.Drivers[0].LicenseDate = CheckIfNull(excelAcess, i, 13);
            dataModel.policy.Drivers[0].LicenseIssueDate = excelAcess.xlRange.Cells[i, 14].Value2;
            dataModel.policy.Drivers[0].Marital = excelAcess.xlRange.Cells[i, 15].Value2.ToString();
            dataModel.policy.Drivers[0].Name = excelAcess.xlRange.Cells[i, 16].Value2;
            dataModel.policy.EffectiveDate = excelAcess.xlRange.Cells[i, 17].Value2;
            dataModel.policy.EffectiveTime = excelAcess.xlRange.Cells[i, 18].Value2;
            dataModel.policy.ExpirationDate = excelAcess.xlRange.Cells[i, 19].Value2;
            dataModel.policy.ExpirationTime = excelAcess.xlRange.Cells[i, 20].Value2;
            dataModel.policy.Insured.Address.Appartment = CheckIfNull(excelAcess, i, 23);
            dataModel.policy.Insured.Address.Building = CheckIfNull(excelAcess, i, 24);
            dataModel.policy.Insured.Address.City = CheckIfNull(excelAcess, i, 25);
            dataModel.policy.Insured.Address.FullAddress = excelAcess.xlRange.Cells[i, 26].Value2;
            dataModel.policy.Insured.Address.GNINMB = CheckIfNull(excelAcess, i, 27);
            dataModel.policy.Insured.Address.House = excelAcess.xlRange.Cells[i, 28].Value2.ToString();
            dataModel.policy.Insured.Address.KladrCode = excelAcess.xlRange.Cells[i, 29].Value2;
            dataModel.policy.Insured.Address.RegionArea = CheckIfNull(excelAcess, i, 30);
            dataModel.policy.Insured.Address.Street = excelAcess.xlRange.Cells[i, 31].Value2;
            dataModel.policy.Insured.Address.Subject = excelAcess.xlRange.Cells[i, 32].Value2;
            dataModel.policy.Insured.Address.Town = CheckIfNull(excelAcess, i, 33);
            dataModel.policy.Insured.Address.ZIP = excelAcess.xlRange.Cells[i, 34].Value2.ToString();
            dataModel.policy.Insured.Birthdate = excelAcess.xlRange.Cells[i, 35].Value2;
            dataModel.policy.Insured.FirstName = excelAcess.xlRange.Cells[i, 36].Value2;
            dataModel.policy.Insured.SecondName = excelAcess.xlRange.Cells[i, 37].Value2;
            dataModel.policy.Insured.ThirdName = excelAcess.xlRange.Cells[i, 38].Value2;
            dataModel.policy.Insured.Type = excelAcess.xlRange.Cells[i, 39].Value2;
            dataModel.policy.InsuredName = excelAcess.xlRange.Cells[i, 40].Value2;
            dataModel.policy.IsProlongation = StringToBool(excelAcess.xlRange.Cells[i, 41].Value2);
            dataModel.policy.Parameters.Duration = Convert.ToInt32(excelAcess.xlRange.Cells[i, 43].Value2);
            dataModel.policy.PaymentCurrency = excelAcess.xlRange.Cells[i, 44].Value2;
            dataModel.policy.PriorPolicyNumber = CheckIfNull(excelAcess, i, 45);
            dataModel.policy.PriorPolicySerial = CheckIfNull(excelAcess, i, 46);
            dataModel.policy.ProductID = excelAcess.xlRange.Cells[i, 47].Value2;
            dataModel.policy.ProductName = excelAcess.xlRange.Cells[i, 48].Value2;
            dataModel.policy.Accident.Premium = Convert.ToInt32(excelAcess.xlRange.Cells[i, 50].Value2);
            dataModel.policy.Accident.Sum = Convert.ToInt32(excelAcess.xlRange.Cells[i, 51].Value2);
            dataModel.policy.Risks.TypeOne.Selected = StringToBool(excelAcess.xlRange.Cells[i, 54].Value2);
            dataModel.policy.Risks.TypeOne.Sum = excelAcess.xlRange.Cells[i, 55].Value2;
            dataModel.policy.Risks.TypeTwo.Selected = StringToBool(excelAcess.xlRange.Cells[i, 57].Value2);
            dataModel.policy.Risks.TypeTwo.Sum = excelAcess.xlRange.Cells[i, 58].Value2;
            dataModel.policy.Risks.TypeThree.Selected = StringToBool(excelAcess.xlRange.Cells[i, 60].Value2);
            dataModel.policy.Risks.TypeThree.Sum = excelAcess.xlRange.Cells[i, 61].Value2;
            dataModel.policy.Risks.TypeFour.Selected = StringToBool(excelAcess.xlRange.Cells[i, 63].Value2);
            dataModel.policy.Risks.TypeFour.Sum = excelAcess.xlRange.Cells[i, 64].Value2;
            dataModel.policy.StatusID = excelAcess.xlRange.Cells[i, 65].Value2;
            dataModel.policy.Vehicle.Age = Convert.ToInt32(excelAcess.xlRange.Cells[i, 67].Value2);
            dataModel.policy.Vehicle.AutoStart = excelAcess.xlRange.Cells[i, 68].Value2;
            dataModel.policy.Vehicle.CreditBankId = CheckIfNull(excelAcess, i, 69);
            dataModel.policy.Vehicle.CreditBankText = CheckIfNull(excelAcess, i, 70);
            dataModel.policy.Vehicle.EngineVol = excelAcess.xlRange.Cells[i, 71].Value2.ToString();
            dataModel.policy.Vehicle.HorsePower = excelAcess.xlRange.Cells[i, 72].Value2.ToString();
            dataModel.policy.Vehicle.IsCreditCar = StringToBool(excelAcess.xlRange.Cells[i, 73].Value2);
            dataModel.policy.Vehicle.IsNew = StringToBool(excelAcess.xlRange.Cells[i, 74].Value2);
            dataModel.policy.Vehicle.IsUsed = StringToBool(excelAcess.xlRange.Cells[i, 75].Value2);
            dataModel.policy.Vehicle.IsUsedTransfer = StringToBool(excelAcess.xlRange.Cells[i, 76].Value2);
            dataModel.policy.Vehicle.MarkId = excelAcess.xlRange.Cells[i, 77].Value2;
            dataModel.policy.Vehicle.MarkText = excelAcess.xlRange.Cells[i, 78].Value2;
            dataModel.policy.Vehicle.Mileage = excelAcess.xlRange.Cells[i, 79].Value2.ToString();
            dataModel.policy.Vehicle.ModelId = excelAcess.xlRange.Cells[i, 80].Value2;
            dataModel.policy.Vehicle.ModelText = excelAcess.xlRange.Cells[i, 81].Value2.ToString();
            dataModel.policy.Vehicle.Price = excelAcess.xlRange.Cells[i, 82].Value2.ToString();
            dataModel.policy.Vehicle.Righthand = excelAcess.xlRange.Cells[i, 83].Value2;
            dataModel.policy.Vehicle.TypeId = excelAcess.xlRange.Cells[i, 84].Value2;
            dataModel.policy.Vehicle.TypeText = excelAcess.xlRange.Cells[i, 85].Value2;
            dataModel.policy.Vehicle.Year = Convert.ToInt32(excelAcess.xlRange.Cells[i, 86].Value2);
            dataModel.policy.Vehicle.UsagePurposeId = excelAcess.xlRange.Cells[i, 87].Value2;
            dataModel.policy.Vehicle.UsagePurposeText = excelAcess.xlRange.Cells[i, 88].Value2;
            dataModel.policy.Vehicle.Chassis = CheckIfNull(excelAcess, i, 89);
            dataModel.policy.Beneficiary.BeneficiaryId = excelAcess.xlRange.Cells[i, 92].Value2.ToString();
            dataModel.policy.Beneficiary.BeneficiaryType = CheckIfNull(excelAcess, i, 93);
        }

        public bool StringToBool(string value)
        {
            switch (value)
            {
                case "true":
                    return true;
                    break;
                case "false":
                    return false;
                    break;
                default:
                    return false;
                    break;
            }
        }
    }
}
