using System;

namespace RPAUserAdminOJT.Controllers.Models
{
    public class AdUser
    {
        public string Description { get; set; }
        public string DisplayName { get; set; }
        public string DistinguishedName { get; set; }
        public string EmailAddress { get; set; }
        public string EmployeeId { get; set; }
        public string GivenName { get; set; }
        public string MiddleName { get; set; }
        public string Name { get; set; }
        public bool PasswordNeverExpires { get; set; }
        public bool PasswordNotRequired { get; set; }
        public bool Enable { set; get; }
        public string SamAccountName { get; set; }
        public string FirsLastName { set; get; }
        public string SecondLastName { set; get; }
        public bool UserCannotChangePassword { get; set; }
        private string UserPrincipalName { get; set; }
        public string HomePage { set; get; }
        public string Domain { set; get; }
        public string PostOfficeBox { set; get; }
        public string City { set; get; }
        public string State { set; get; }
        public string Country { set; get; }
        public string Company { set; get; }
        public string Department { set; get; }
        public string Password { set; get; }


        public string Surname()
        {
            if (String.IsNullOrEmpty(SecondLastName))
            {
                return Utility.TextFormatter.FirstLetterToUpperCase(FirsLastName.ToLower());
            }
            return $"{ Utility.TextFormatter.FirstLetterToUpperCase(FirsLastName.ToLower()) } { Utility.TextFormatter.FirstLetterToUpperCase(SecondLastName.ToLower()) }";
        }

        public string GetFullName()
        {
            return $"{ GetName() } {Surname()}";
        }

        public string GetUserPrincipalName()
        {
            return $"{SamAccountName}@{Domain}";
        }

        public string GetHomePage()
        {
            return $"{SamAccountName}@{HomePage}";
        }

        public string GetName()
        {
            if (String.IsNullOrEmpty(MiddleName))
            {
                return Utility.TextFormatter.FirstLetterToUpperCase(GivenName.ToLower());
            }
            return $"{ Utility.TextFormatter.FirstLetterToUpperCase(GivenName.ToLower())} { Utility.TextFormatter.FirstLetterToUpperCase(MiddleName.ToLower())}";
        }
    }
}