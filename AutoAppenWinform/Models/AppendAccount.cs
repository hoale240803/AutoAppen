using AutoAppenWinform.Enum;

namespace AutoAppenWinform.Models
{
    public class AppendAccount
    {
        public string? Email { get; set; }
        public string Password { get; set; }
        public string FirstName { get; set; }

        public string LastName { get; set; }
        public CountryEnum CountryOfResidence { get; set; }
    }
}