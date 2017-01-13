using AspNet.Owin.SharePoint.Addin.Authentication;
using AspNet.Owin.SharePoint.Addin.Authentication.Middleware;
using Microsoft.Owin.Security;
using SPMVCWeb.Helpers;
using System;
using System.Configuration;
using System.Net.Http;
using System.Security.Claims;
using System.Web;
using System.Web.Configuration;
using System.Web.Mvc;

namespace SPMVCWeb.Controllers
{
    [AllowAnonymous]
    public class AccountController : Controller
    {
        private Uri GetSPHostUrl(string url)
        {
            var queryString = new Uri(string.Concat("http://tempuri.org", url)).ParseQueryString();
            string spHostUrlString = TokenHelper.EnsureTrailingSlash(queryString[SharePointContext.SPHostUrlKey]);
            if (string.IsNullOrEmpty(spHostUrlString))
            {
                spHostUrlString = ConfigurationManager.AppSettings[SharePointContext.SPHostUrlKey];
            }
            Uri spHostUrl;
            if (Uri.TryCreate(spHostUrlString, UriKind.Absolute, out spHostUrl) &&
                (spHostUrl.Scheme == Uri.UriSchemeHttp || spHostUrl.Scheme == Uri.UriSchemeHttps))
            {
                return spHostUrl;
            }
            return null;
        }

        [HttpGet]
        public ActionResult Login(string returnUrl)
        {
            if (string.IsNullOrEmpty(returnUrl))
                returnUrl = "/";

            var cookieAuthenticationEnabled = string.IsNullOrEmpty(WebConfigurationManager.AppSettings.Get("CookieAuthenticationEnabled")) ? false : Convert.ToBoolean(WebConfigurationManager.AppSettings.Get("CookieAuthenticationEnabled"));
            if (cookieAuthenticationEnabled)
            {
                Uri spHostUrl = GetSPHostUrl(returnUrl);
                if (spHostUrl == null)
                {
                    throw new Exception(string.Format("Unable to determine {0}.", SharePointContext.SPHostUrlKey));
                }
                return new ChallengeResult(SPAddinAuthenticationDefaults.AuthenticationType, spHostUrl.ToString(), returnUrl);
            }
            Uri redirectUrl;
            Uri returnUri;
            if (!Uri.TryCreate(returnUrl, UriKind.RelativeOrAbsolute, out returnUri))
            {
                returnUri = null;
            }
            else
            {
                if (!returnUri.IsAbsoluteUri)
                {
                    Uri.TryCreate(HttpContext.Request.Url, returnUrl, out returnUri);
                }
            }
            switch (SPContextHelper.CheckRedirectionStatus(HttpContext, returnUri, out redirectUrl))
            {
                case RedirectionStatus.Ok:
                    return new RedirectResult(returnUrl);
                case RedirectionStatus.ShouldRedirect:
                    return new RedirectResult(redirectUrl.AbsoluteUri);
                case RedirectionStatus.CanNotRedirect:
                    return new ViewResult { ViewName = "Error" };
            }
            return new RedirectResult(returnUrl);
        }

        [HttpGet]
        public ActionResult Logout(string returnUrl)
        {
            var cookieAuthenticationEnabled = string.IsNullOrEmpty(WebConfigurationManager.AppSettings.Get("CookieAuthenticationEnabled")) ? false : Convert.ToBoolean(WebConfigurationManager.AppSettings.Get("CookieAuthenticationEnabled"));
            if (cookieAuthenticationEnabled)
            {
                if (Request.IsAuthenticated)
                {
                    SPContext spContext = SPContextProvider.Get(HttpContext.User as ClaimsPrincipal);
                    spContext.ClearCache();
                    HttpContext.GetOwinContext().Authentication.SignOut(SPAddinAuthenticationDefaults.AuthenticationType);
                    if (spContext.SPAppWebUrl != null)
                    {
                        return new RedirectResult(string.Format("{0}/_layouts/closeConnection.aspx?loginasanotheruser=true",
                                spContext.SPAppWebUrl.GetLeftPart(UriPartial.Path).TrimEnd('/')));
                    }
                }
            }
            else
            {
                Uri spHostUrl = GetSPHostUrl(returnUrl);
                if (spHostUrl == null)
                {
                    spHostUrl = SharePointContext.GetSPHostUrl(HttpContext.Request);
                }
                var spContext = SharePointContextProvider.Current.GetSharePointContext(HttpContext, spHostUrl);
                if (spContext != null)
                {
                    Uri appWebUrl = spContext.SPAppWebUrl;
                    SharePointContextProvider.Current.ClearCache(HttpContext);
                    if (appWebUrl != null)
                    {
                        return new RedirectResult(string.Format("{0}/_layouts/closeConnection.aspx?loginasanotheruser=true",
                                appWebUrl.GetLeftPart(UriPartial.Path).TrimEnd('/')));
                    }
                }
            }
            return new RedirectResult(string.Format("/login?ReturnUrl={0}", HttpUtility.UrlEncode(returnUrl)));
        }

        private class ChallengeResult : HttpUnauthorizedResult
        {
            public ChallengeResult(string provider, string spHostUrl, string redirectUri)
            {
                LoginProvider = provider;
                RedirectUri = redirectUri;
                SPHostUrl = spHostUrl;
            }

            private string LoginProvider { get; }
            private string RedirectUri { get; }
            private string SPHostUrl { get; }

            public override void ExecuteResult(ControllerContext context)
            {
                var properties = new AuthenticationProperties { RedirectUri = RedirectUri };
                properties.Dictionary[SharePointContext.SPHostUrlKey] = SPHostUrl;
                context.HttpContext.GetOwinContext().Authentication.Challenge(properties, LoginProvider);
            }
        }
    }
}