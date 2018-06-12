using System.Web;
using System.Web.Mvc;
using Microsoft.Owin.Security;
using Microsoft.Owin.Security.Cookies;
using Microsoft.Owin.Security.OpenIdConnect;
using System.Security.Claims;
using MSGraphCalendarViewer.TokenStorage;
using MSGraphCalendarViewer.Helpers;

namespace MSGraphCalendarViewer.Controllers
{
  public class AccountController : Controller
  {
    public void SignIn()
    {
      if (!Request.IsAuthenticated)
      {
        HttpContext.GetOwinContext().Authentication.Challenge(
          new AuthenticationProperties { RedirectUri = "/" },
          OpenIdConnectAuthenticationDefaults.AuthenticationType);
      }
    }

    public void SignOut()
    {
      if (Request.IsAuthenticated)
      {
        string userObjectId = ClaimsPrincipal.Current.FindFirst(ClaimTypes.NameIdentifier).Value;

        SessionTokenCache tokenCache = new SessionTokenCache(userObjectId, HttpContext);
        HttpContext.GetOwinContext().Authentication.SignOut(OpenIdConnectAuthenticationDefaults.AuthenticationType, CookieAuthenticationDefaults.AuthenticationType);
      }

      HttpContext.GetOwinContext().Authentication.SignOut(CookieAuthenticationDefaults.AuthenticationType);
      Response.Redirect("/");
    }
  }
}