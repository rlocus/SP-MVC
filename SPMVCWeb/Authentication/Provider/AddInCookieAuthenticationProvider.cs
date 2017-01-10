using AspNet.Owin.SharePoint.Addin.Authentication.Common;
using AspNet.Owin.SharePoint.Addin.Authentication.Context;
using Microsoft.Owin.Security.Cookies;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Claims;
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
                if (context.Request.Path.Value.Contains(context.Options.LoginPath.Value))
                {
                    return Task.FromResult<object>(null);
                }
                var spContext = SPContextProvider.Get(new ClaimsPrincipal(context.Identity));
                string spHostUrlString = TokenHelper.EnsureTrailingSlash(context.Request.Query.Get(SharePointContext.SPHostUrlKey));
                Uri spHostUrl;
                if (!Uri.TryCreate(spHostUrlString, UriKind.Absolute, out spHostUrl))
                {
                    throw new Exception(string.Format("Unable to determine {0}.", SharePointContext.SPHostUrlKey));
                }
                //string spAppWebUrlString = TokenHelper.EnsureTrailingSlash(context.Request.Query.Get(SharePointContext.SPAppWebUrlKey));
                //Uri spAppWebUrl;
                //if (string.IsNullOrEmpty(spAppWebUrlString) || (!Uri.TryCreate(spAppWebUrlString, UriKind.Absolute, out spAppWebUrl) &&
                //                                (spAppWebUrl.Scheme == Uri.UriSchemeHttp || spAppWebUrl.Scheme == Uri.UriSchemeHttps)))
                //{
                //    throw new Exception(string.Format("Unable to determine {0}.", SharePointContext.SPAppWebUrlKey));
                //}
                if (spContext != null)
                {
                    if (spHostUrl != null && !spContext.SPHostUrl.AbsoluteUri.TrimEnd('/').Equals(spHostUrl.AbsoluteUri.TrimEnd('/'), StringComparison.OrdinalIgnoreCase))
                    {
                        context.RejectIdentity();
                    }
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