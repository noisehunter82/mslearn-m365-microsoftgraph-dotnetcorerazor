using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using DotNetCoreRazor_MSGraph.Graph;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.Extensions.Logging;
using Microsoft.Graph;
using Microsoft.Identity.Web;

namespace DotNetCoreRazor_MSGraph.Pages
{
    [AuthorizeForScopes(ScopeKeySection = "DownstreamApi:Scopes")]
    public class EmailModel : PageModel
    {
        private readonly GraphEmailClient _graphEmailClient;

        [BindProperty(SupportsGet = true)]
        public string NextLink { get; set; }

        public string CurrentLink { get; set; }

        public string PreviousLink { get; set; }

        public int SkipNumber { get; set; }

        public IEnumerable<Message> Messages { get; private set; }

        public EmailModel(GraphEmailClient graphEmailClient)
        {
            _graphEmailClient = graphEmailClient;
        }

        public async Task OnGetAsync()
        {
            var messagesPagingData = await _graphEmailClient.GetUserMessagesPage(NextLink);
            Messages = messagesPagingData.Messages;
            NextLink = messagesPagingData.NextLink;
            CurrentLink = messagesPagingData.CurrentLink;
            if (!string.IsNullOrEmpty(CurrentLink))
            {
                SkipNumber = messagesPagingData.SkipNumber;
                var skipValue = int.Parse(CurrentLink.Split("skip=")[1]);
                if (skipValue >= SkipNumber)
                {
                    PreviousLink = CurrentLink.Replace($"skip={skipValue}", $"skip={skipValue - SkipNumber}");
                }
            }
            await Task.CompletedTask;
        }
    }
}
