using System.Configuration;
using System.IdentityModel.Claims;
using System.Threading.Tasks;
using System.Web;
using Owin;
using MSGraphCalendarViewer.TokenStorage;
using Microsoft.Identity.Client;
using Microsoft.IdentityModel.Tokens;
using Microsoft.Owin.Security;
using Microsoft.Owin.Security.Cookies;
using Microsoft.Owin.Security.OpenIdConnect;

namespace MSGraphCalendarViewer
{
  public partial class Startup
  {
    // The appId is used by the application to uniquely identify itself to Azure AD.
    // The appSecret is the application's password.
    // The redirectUri is where users are redirected after sign in and consent.
    // The graphScopes are the Microsoft Graph permission scopes that are used by this sample: User.Read Mail.Send
    private static string appId = ConfigurationManager.AppSettings["ida:AppId"];
    private static string appSecret = ConfigurationManager.AppSettings["ida:AppSecret"];
    private static string redirectUri = ConfigurationManager.AppSettings["ida:RedirectUri"];
    private static string graphScopes = ConfigurationManager.AppSettings["ida:GraphScopes"];

    public void ConfigureAuth(IAppBuilder app)
    {
      app.SetDefaultSignInAsAuthenticationType(CookieAuthenticationDefaults.AuthenticationType);

      app.UseCookieAuthentication(new CookieAuthenticationOptions());

      app.UseOpenIdConnectAuthentication(
          new OpenIdConnectAuthenticationOptions
          {
            ClientId = appId,
            Authority = "https://login.microsoftonline.com/common/v2.0",
            PostLogoutRedirectUri = redirectUri,
            RedirectUri = redirectUri,
            Scope = "openid email profile offline_access " + graphScopes,
            TokenValidationParameters = new TokenValidationParameters
            {
              ValidateIssuer = false,
            },
            Notifications = new OpenIdConnectAuthenticationNotifications
            {
              AuthorizationCodeReceived = async (context) =>
              {
                var code = context.Code;
                string signedInUserID = context.AuthenticationTicket.Identity.FindFirst(ClaimTypes.NameIdentifier).Value;

                TokenCache userTokenCache = new SessionTokenCache(
                  signedInUserID,
                  context.OwinContext.Environment["System.Web.HttpContextBase"] as HttpContextBase
                ).GetMsalCacheInstance();
                ConfidentialClientApplication cca = new ConfidentialClientApplication(
                  appId,
                  redirectUri,
                  new ClientCredential(appSecret),
                  userTokenCache,
                  null);

                string[] scopes = graphScopes.Split(new char[] { ' ' });
                AuthenticationResult result = await cca.AcquireTokenByAuthorizationCodeAsync(code, scopes);
              },
              AuthenticationFailed = (context) =>
              {
                context.HandleResponse();
                context.Response.Redirect("/Error?message=" + context.Exception.Message);
                return Task.FromResult(0);
              }
            }
          }
      );

    }
  }
}