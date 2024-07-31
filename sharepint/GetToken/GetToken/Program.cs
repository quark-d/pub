// See https://aka.ms/new-console-template for more information



using System;
using System.Threading.Tasks;
using Microsoft.Identity.Client;
using TextCopy;

namespace AzureADToken
{
  class Program
  {
    private const string clientId = "";
    private const string tenantId = "";
    private const string clientSecret = "";
    private const string authority = $"https://login.microsoftonline.com/{tenantId}";
    private const string scope = "https://graph.microsoft.com/.default";

    static async Task Main(string[] args)
    {
      var app = ConfidentialClientApplicationBuilder.Create(clientId)
          .WithClientSecret(clientSecret)
          .WithAuthority(new Uri(authority))
          .Build();

      var scopes = new string[] { scope };

      try
      {
        Console.WriteLine("開始");

        var result = await app.AcquireTokenForClient(scopes)
            .ExecuteAsync();

        string accessToken = result.AccessToken;
        Console.WriteLine("Access Token: " + accessToken);

        // トークンをクリップボードに設定
        ClipboardService.SetText(accessToken);
        Console.WriteLine("Access Token has been copied to the clipboard.");
      }
      catch (Exception ex)
      {
        Console.WriteLine("Error acquiring token: " + ex.Message);
      }
    }
  }
}
