using AspNet.Owin.SharePoint.Addin.Authentication.Common;
using AspNet.Owin.SharePoint.Addin.Authentication.Context;
using Microsoft.Owin.Security.Cookies;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Web;

namespace AspNet.Owin.SharePoint.Addin.Authentication.Provider
{
    public class AddInCookieAuthenticationProvider : ICookieAuthenticationProvider
    {
        public Task ValidateIdentity(CookieValidateIdentityContext context)
        {
            if (context.Identity.IsAuthenticated)
            {
                //var queryStringHostUrl = context.Request.Query["h"];
               
                    if (context.Request.Path.Value.Contains("Auth") ||
                    context.Request.Path.Value.StartsWith("signin") || context.Request.Path.Value.Contains(context.Options.LoginPath.Value))
                {
                    return Task.FromResult<object>(null);
                }

                Uri spHostUrl;
                if (!Uri.TryCreate(context.Request.Query[SharePointContext.SPHostUrlKey], UriKind.Absolute, out spHostUrl))
                {
                    throw new Exception("Unable to determine host url");
                }

                var hostUrl = context.Identity.FindFirst(SPAddinClaimTypes.SPHostUrl).Value;

                if (!hostUrl.Equals(spHostUrl.AbsoluteUri, StringComparison.OrdinalIgnoreCase))
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