using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Util.Store;
using System;
using System.IO;
using System.Threading;

public class GoogleService
{

    private readonly string _googleSecretJsonFilePath;
    private readonly string _applicationName;
    private readonly string[] _scopes;

    public GoogleService(string googleSecretJsonFilePath, string applicationName, string[] scopes)
    {
        _googleSecretJsonFilePath = googleSecretJsonFilePath;
        _applicationName = applicationName;
        _scopes = scopes;
    }

    public UserCredential GetGoogleCredential()
    {
  

        UserCredential credential;


        using (var stream =
               new FileStream(_googleSecretJsonFilePath, FileMode.Open, FileAccess.Read))
        {
            string credPath = "token.json";
            credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                   GoogleClientSecrets.Load(stream).Secrets,
                   _scopes,
                   "user",
                   CancellationToken.None,
                   new FileDataStore(credPath, true)).Result;
          
        }
        return credential;

    }


     
    

    public SheetsService GetSheetsService()
    {
        var credential = GetGoogleCredential();
        var sheetsService = new SheetsService(new BaseClientService.Initializer()
        {
            HttpClientInitializer = credential,
            ApplicationName = _applicationName,
        });
        return sheetsService;
    }
}
