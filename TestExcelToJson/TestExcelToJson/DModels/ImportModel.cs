using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TestExcelToJson.DModels
{
    public class ImportModel
    {

        public class Rootobject
        {
            public Model model { get; set; }
        }

        public class Model
        {
            public Deal[] deals { get; set; }
            public Dealsobject[] dealsObjects { get; set; }
        }

        public class Deal
        {
            public string dealType { get; set; }
            public Insuranceobject[] insuranceObjects { get; set; }
            public Insurant insurant { get; set; }
        }

        public class Insurant
        {
            public int dealObjects { get; set; }
        }

        public class Insuranceobject
        {
            public Auto auto { get; set; }
            public Owner[] owners { get; set; }
        }

        public class Auto
        {
            public int dealObjects { get; set; }
        }

        public class Owner
        {
            public int dealObjects { get; set; }
        }

        public class Dealsobject
        {
            public int id { get; set; }
            public string type { get; set; }
            public Person person { get; set; }
            public Auto1 auto { get; set; }
        }

        public class Person
        {
            public string dateOfBirth { get; set; }
            public string email1 { get; set; }
            public Identitydocument identityDocument { get; set; }
            public bool hasPersonalDataAgreement { get; set; }
            public string name { get; set; }
            public string lastName { get; set; }
            public string middleName { get; set; }
            public string phone1 { get; set; }
        }

        public class Identitydocument
        {
            public Registrationaddress registrationAddress { get; set; }
            public string dateIssue { get; set; }
            public string number { get; set; }
            public string series { get; set; }
            public string identityDocumentTypeCode { get; set; }
            public string citizenshipCode { get; set; }
        }

        public class Registrationaddress
        {
            public string country { get; set; }
            public string region { get; set; }
            public string city { get; set; }
            public string district { get; set; }
            public string street { get; set; }
            public string house { get; set; }
            public string fiasId { get; set; }
            public string kladrId { get; set; }
        }

        public class Auto1
        {
            public string registrationNumber { get; set; }
            public Tsidentitycard tsIdentityCard { get; set; }
            public string vin { get; set; }
        }

        public class Tsidentitycard
        {
            public string dateOfIssued { get; set; }
            public string number { get; set; }
            public string series { get; set; }
        }

    }
}
