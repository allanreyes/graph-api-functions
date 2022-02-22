using CsvHelper;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using Microsoft.Graph;
using Microsoft.Graph.Auth;
using Microsoft.Identity.Client;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace GraphAPIFunctions
{
    public class GetNewUsers
    {
        private readonly IConfiguration _config;
        private readonly IConfidentialClientApplication _confidentialClientApplication;

        public GetNewUsers(IConfiguration config)
        {
            _config = config;
            _confidentialClientApplication = ConfidentialClientApplicationBuilder
                  .Create(_config["ClientId"])
                  .WithTenantId(_config["TenantId"])
                  .WithClientSecret(_config["ClientSecret"])
                  .Build();
        }

        [FunctionName(nameof(GetNewUsers))]
        public async Task<IActionResult> Run([HttpTrigger(AuthorizationLevel.Function, "get", Route = null)] HttpRequest req, ILogger log)
        {
            var result = new List<MailboxUsageDetails>();
            var payload = req.Query["dateFrom"];
            var dateFrom = DateTime.ParseExact(payload, "yyyy-MM-dd", CultureInfo.InvariantCulture);

            var graphClient = new GraphServiceClient(new ClientCredentialProvider(_confidentialClientApplication));

            try
            {
                var stream = await graphClient.Reports.GetMailboxUsageDetail("D30").Request().GetAsync();
                using (var reader = new StreamReader(stream))
                using (var csv = new CsvReader(reader, CultureInfo.InvariantCulture))
                {
                    result = csv.GetRecords<MailboxUsageDetails>().ToList();
                }
                if (result.Any())
                {
                    result = result.Where(m => m.IsDeleted.Equals("False", StringComparison.InvariantCultureIgnoreCase)
                    && DateTime.ParseExact(m.CreatedDate, "yyyy-MM-dd", CultureInfo.InvariantCulture) >= dateFrom
                    ).ToList();
                }
            }
            catch (Exception ex)
            {
                return new BadRequestObjectResult(ex);
            }


            return new OkObjectResult(result);
        }


    }
}
