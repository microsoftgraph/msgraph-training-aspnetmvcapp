using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Security.Claims;
using System.Threading.Tasks;
using System.Web;
using Microsoft.Identity.Client;
using Microsoft.Owin.Security;
using Microsoft.Owin.Security.OpenIdConnect;
using MSGraphCalendarViewer.TokenStorage;

namespace MSGraphCalendarViewer.Helpers
{
  public sealed class SampleAuthProvider : IAuthProvider
  {
    private string redirectUri = ConfigurationManager.AppSettings["ida:RedirectUri"];
    private string appId = ConfigurationManager.AppSettings["ida:AppId"];
    private string appSecret = ConfigurationManager.AppSettings["ida:AppSecret"];
    private string scopes = ConfigurationManager.AppSettings["ida:GraphScopes"];
    private SessionTokenCache tokenCache { get; set; }

    private static readonly SampleAuthProvider instance = new SampleAuthProvider();
    private SampleAuthProvider() { }

    public static SampleAuthProvider Instance
    {
      get
      {
        return instance;
      }
    }

    public async Task<string> GetUserAccessTokenAsync()
    {
      string signedInUserID = ClaimsPrincipal.Current.FindFirst(ClaimTypes.NameIdentifier).Value;
      HttpContextWrapper httpContext = new HttpContextWrapper(HttpContext.Current);
      TokenCache userTokenCache = new SessionTokenCache(signedInUserID, httpContext).GetMsalCacheInstance();

      ConfidentialClientApplication cca = new ConfidentialClientApplication(
        appId,
        redirectUri,
        new ClientCredential(appSecret),
        userTokenCache,
        null);

      try
      {
        AuthenticationResult result = await cca.AcquireTokenSilentAsync(scopes.Split(new char[] { ' ' }), cca.Users.First());
        return result.AccessToken;
      }

      catch (Exception)
      {
        HttpContext.Current.Request.GetOwinContext().Authentication.Challenge(
          new AuthenticationProperties() { RedirectUri = "/" },
          OpenIdConnectAuthenticationDefaults.AuthenticationType);

        throw new Exception();
      }
    }
  }
}
