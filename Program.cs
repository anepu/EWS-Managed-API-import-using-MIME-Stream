using Microsoft.Identity.Client;
using Microsoft.Exchange.WebServices.Data;
using System;
using System.IO;
using System.Net.Http.Headers;
using System.Threading.Tasks;

class Program
{
    // Declare the mailbox to impersonate
    private const string mailbox = "user@contoso.com";

    private static async System.Threading.Tasks.Task Main(string[] args)
    {
        var service = await AuthenticateExchangeService();
        string targetMailbox = mailbox; 
        UploadMIMEEmail(service, targetMailbox);
    }

    private static async Task<ExchangeService> AuthenticateExchangeService()
    {
        // Declare your app registration details
        string clientId = "890af5cc-1111-4c16-b5a8-53d7a3465064";
        string tenantId = "fb97d94b-1111-457d-927b-37db26d21a2c";
        string clientSecret = "secretkey";
        string impersonatedAccount = mailbox;


        var app = ConfidentialClientApplicationBuilder.Create(clientId)
           .WithClientSecret(clientSecret)
           .WithAuthority(new Uri($"https://login.microsoftonline.com/{tenantId}"))
           .Build();

        string[] scopes = new string[] { "https://outlook.office365.com/.default" };
        AuthenticationResult result = await app.AcquireTokenForClient(scopes).ExecuteAsync();

        var service = new ExchangeService
        {
            Url = new Uri("https://outlook.office365.com/EWS/Exchange.asmx"),
            Credentials = new OAuthCredentials(result.AccessToken)
        };

        // Set Exchange Impersonation
        service.ImpersonatedUserId = new ImpersonatedUserId(ConnectingIdType.SmtpAddress, impersonatedAccount);

        return service;
    }

    private static void UploadMIMEEmail(ExchangeService service, string targetMailbox)
    {
        EmailMessage email = new EmailMessage(service);

        string emlFileName = @"C:\temp\test3.eml";
        using (FileStream fs = new FileStream(emlFileName, FileMode.Open, FileAccess.Read))
        {
            byte[] bytes = new byte[fs.Length];
            int numBytesToRead = (int)fs.Length;
            int numBytesRead = 0;
            while (numBytesToRead > 0)
            {
                int n = fs.Read(bytes, numBytesRead, numBytesToRead);
                if (n == 0)
                    break;
                numBytesRead += n;
                numBytesToRead -= n;
            }
            // Set the contents of the .eml file to the MimeContent property.
            email.MimeContent = new MimeContent("UTF-8", bytes);
        }

        // Indicate that this email is not a draft. Otherwise, the email will appear as a 
        // draft to clients.
        ExtendedPropertyDefinition PR_MESSAGE_FLAGS_msgflag_read = new ExtendedPropertyDefinition(3591, MapiPropertyType.Integer);
        email.SetExtendedProperty(PR_MESSAGE_FLAGS_msgflag_read, 1);

        // Specify the target folder in the specified mailbox
        FolderId targetFolderId = new FolderId(WellKnownFolderName.Inbox, targetMailbox);

        // Save the email to the specified folder
        email.Save(targetFolderId);
    }
}
