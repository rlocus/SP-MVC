using Microsoft.Owin.Security.Cookies;
using System;
using System.Configuration;
using System.Security.Claims;
using System.Threading.Tasks;

namespace AspNet.Owin.SharePoint.Addin.Authentication.Provider
{
  public class AddInCookieAuthenticationProvider : ICookieAuthenticationProvider
  {
    public Task ValidateIdentity(CookieValidateIdentityContext context)
    {
      if (context.Identity.IsAuthenticated)
      {
        if (context.Request.Path.Value.Contains(context.Options.LoginPath.Value))
        {
          return Task.FromResult<object>(null);
        }
        bool isWebPart = context.Request.Get<string>("IsWebPart") == "1";
        var spContext = SPContextProvider.Get(context.Identity, isWebPart);
        string spHostUrlString = TokenHelper.EnsureTrailingSlash(context.Request.Query.Get(SharePointContext.SPHostUrlKey));
        if (string.IsNullOrEmpty(spHostUrlString))
        {
          spHostUrlString = ConfigurationManager.AppSettings["SPHostUrl"];
        }
        Uri spHostUrl;
        if (!Uri.TryCreate(spHostUrlString, UriKind.Absolute, out spHostUrl))
        {
          //throw new Exception(string.Format("Unable to determine {0}.", SharePointContext.SPHostUrlKey));
        }

        //try
        //{
        if (spHostUrl != null &&
            !string.Equals(spContext.SPHostUrl.GetLeftPart(UriPartial.Path).TrimEnd('/'),
                spHostUrl.GetLeftPart(UriPartial.Path).TrimEnd('/'), StringComparison.OrdinalIgnoreCase))
        {
          context.RejectIdentity();
        }
        //}
        //catch (Exception)
        //{
        //    context.RejectIdentity();
        //}

        string clientId = ConfigurationManager.AppSettings["ClientId"];
        try
        {
          if (spContext.ClientId != (string.IsNullOrEmpty(clientId) ? Guid.Empty : new Guid(clientId)))
          {
            context.RejectIdentity();
          }
        }
        catch (Exception)
        {
          context.RejectIdentity();
        }

      }
      return Task.FromResult<object>(null);
    }

    public void ResponseSignIn(CookieResponseSignInContext context)
    {
    }

    public void ApplyRedirect(CookieApplyRedirectContext context)
    {
      context.Response.Redirect(context.RedirectUri);
    }

    public void ResponseSignOut(CookieResponseSignOutContext context)
    {
    }

    public void Exception(CookieExceptionContext context)
    {
    }

    public void ResponseSignedIn(CookieResponseSignedInContext context)
    {
    }
  }
}