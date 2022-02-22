using CsvHelper.Configuration.Attributes;
using Newtonsoft.Json;

namespace GraphAPIFunctions
{
    public class MailboxUsageDetails
    {
        [Index(1)]
        public string UserPrincipalName { get; set; }
        [Index(3)]
        [JsonIgnore]
        public string IsDeleted { get; set; }
        [Index(5)]
        public string CreatedDate { get; set; }
    }
}
