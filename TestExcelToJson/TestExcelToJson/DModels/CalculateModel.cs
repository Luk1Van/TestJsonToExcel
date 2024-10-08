using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TestExcelToJson.DModels
{
    public class CalculateModel
    {

        public  class Rootobject 
        {
            public Policy policy { get; set; }
        }

        public  class Policy 
        {
            public string AmountCurrency { get; set; }
            public string AmountCurrencyName { get; set; }
            public string BSOType { get; set; }
            public string DocumentDate { get; set; }
            public Driver[] Drivers { get; set; }
            public string EffectiveDate { get; set; }
            public string EffectiveTime { get; set; }
            public string ExpirationDate { get; set; }
            public string ExpirationTime { get; set; }
            public Insured Insured { get; set; }
            public string InsuredName { get; set; }
            public bool IsProlongation { get; set; }
            public Parameters Parameters { get; set; }
            public string PaymentCurrency { get; set; }
            public string PriorPolicyNumber { get; set; }
            public string PriorPolicySerial { get; set; }
            public string ProductID { get; set; }
            public string ProductName { get; set; }
            public Accident Accident { get; set; }
            public Risks Risks { get; set; }
            public string StatusID { get; set; }
            public Vehicle Vehicle { get; set; }
            public Paymentdocument PaymentDocument { get; set; }
            public Beneficiary Beneficiary { get; set; }
        }

        public class Insured
        {
            public Address Address { get; set; }
            public string Birthdate { get; set; }
            public string FirstName { get; set; }
            public string SecondName { get; set; }
            public string ThirdName { get; set; }
            public string Type { get; set; }
        }

        public class Address
        {
            public string Appartment { get; set; }
            public string Building { get; set; }
            public string City { get; set; }
            public string FullAddress { get; set; }
            public string GNINMB { get; set; }
            public string House { get; set; }
            public string KladrCode { get; set; }
            public string RegionArea { get; set; }
            public string Street { get; set; }
            public string Subject { get; set; }
            public string Town { get; set; }
            public string ZIP { get; set; }
        }

        public class Parameters
        {
            public int Duration { get; set; }
        }

        public class Accident
        {
            public int Premium { get; set; }
            public int Sum { get; set; }
        }

        public class Risks
        {
            public Typeone TypeOne { get; set; }
            public Typetwo TypeTwo { get; set; }
            public Typethree TypeThree { get; set; }
            public Typefour TypeFour { get; set; }
        }

        public class Typeone
        {
            public bool Selected { get; set; }
            public string Sum { get; set; }
        }

        public class Typetwo
        {
            public bool Selected { get; set; }
            public string Sum { get; set; }
        }

        public class Typethree
        {
            public bool Selected { get; set; }
            public string Sum { get; set; }
        }

        public class Typefour
        {
            public bool Selected { get; set; }
            public string Sum { get; set; }
        }

        public class Vehicle
        {
            public int Age { get; set; }
            public string AutoStart { get; set; }
            public string CreditBankId { get; set; }
            public string CreditBankText { get; set; }
            public string EngineVol { get; set; }
            public string HorsePower { get; set; }
            public bool IsCreditCar { get; set; }
            public bool IsNew { get; set; }
            public bool IsUsed { get; set; }
            public bool IsUsedTransfer { get; set; }
            public string MarkId { get; set; }
            public string MarkText { get; set; }
            public string Mileage { get; set; }
            public string ModelId { get; set; }
            public string ModelText { get; set; }
            public string Price { get; set; }
            public string Righthand { get; set; }
            public string TypeId { get; set; }
            public string TypeText { get; set; }
            public int Year { get; set; }
            public string UsagePurposeId { get; set; }
            public string UsagePurposeText { get; set; }
            public string Chassis { get; set; }
        }

        public class Paymentdocument
        {
        }

        public class Beneficiary
        {
            public string BeneficiaryId { get; set; }
            public string BeneficiaryType { get; set; }
        }

        public class Driver
        {
            public int Age { get; set; }
            public string Birthdate { get; set; }
            public bool DriverInsurer { get; set; }
            public bool DriverOwner { get; set; }
            public int Experience { get; set; }
            public string Gender { get; set; }
            public string LicenseDate { get; set; }
            public string LicenseIssueDate { get; set; }
            public string Marital { get; set; }
            public string Name { get; set; }
        }

    }
}
