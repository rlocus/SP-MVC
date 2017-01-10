using AspNet.Owin.SharePoint.Addin.Authentication.Context;
using AspNet.Owin.SharePoint.Addin.Authentication.Middleware;
using Microsoft.Owin.Security;
using System;
using System.Net.Http;
using System.Web;
using System.Web.Mvc;

namespace SPMVCWeb.Controllers
{
    [AllowAnonymous]
    public class AccountController : Controller
    {
        [HttpGet]
        public ActionResult Login(string returnUrl)
        {
            if (returnUrl == null)
                returnUrl = "/";

            var queryString = new Uri("http://tempuri.org" + returnUrl).ParseQueryString();
            string spHostUrlString = TokenHelper.EnsureTrailingSlash(queryString[SharePointContext.SPHostUrlKey]);
            Uri spHostUrl;
            if (string.IsNullOrEmpty(spHostUrlString) || (!Uri.TryCreate(spHostUrlString, UriKind.Absolute, out spHostUrl) &&
                                            (spHostUrl.Scheme == Uri.UriSchemeHttp || spHostUrl.Scheme == Uri.UriSchemeHttps)))
            {
                throw new Exception(string.Format("Unable to determine {0}.", SharePointContext.SPHostUrlKey));
            }
            //string spAppWebUrlString = TokenHelper.EnsureTrailingSlash(queryString[SharePointContext.SPAppWebUrlKey]);
            //Uri spAppWebUrl;
            //if (string.IsNullOrEmpty(spAppWebUrlString) || (!Uri.TryCreate(spAppWebUrlString, UriKind.Absolute, out spAppWebUrl) &&
            //                                (spAppWebUrl.Scheme == Uri.UriSchemeHttp || spAppWebUrl.Scheme == Uri.UriSchemeHttps)))
            //{
            //    throw new Exception(string.Format("Unable to determine {0}.", SharePointContext.SPAppWebUrlKey));
            //}
            return new ChallengeResult(SPAddinAuthenticationDefaults.AuthenticationType, spHostUrl.ToString(), /*spAppWebUrl.ToString(),*/ returnUrl);
        }

        [HttpPost]
        public void SignOut()
        {
            if (Request.IsAuthenticated)
            {
                HttpContext.GetOwinContext()
                    .Authentication.SignOut(new AuthenticationProperties() { RedirectUri = "/" }, SPAddinAuthenticationDefaults.AuthenticationType);
            }
        }

        public void EndSession()
        {
            HttpContext.GetOwinContext().Authentication.SignOut(SPAddinAuthenticationDefaults.AuthenticationType);
        }

        private class ChallengeResult : HttpUnauthorizedResult
        {
            public ChallengeResult(string provider, string spHostUrl, /*string spAppWebUrl,*/ string redirectUri)
            {
                LoginProvider = provider;
                RedirectUri = redirectUri;
                SPHostUrl = spHostUrl;
                //SPAppWebUrl = spAppWebUrl;
            }

            private string LoginProvider { get; }
            private string RedirectUri { get; }
            private string SPHostUrl { get; }
            //private string SPAppWebUrl { get; }

            public override void ExecuteResult(ControllerContext context)
            {
                var properties = new AuthenticationProperties { RedirectUri = RedirectUri };
                properties.Dictionary[SharePointContext.SPHostUrlKey] = SPHostUrl;
                //properties.Dictionary[SharePointContext.SPAppWebUrlKey] = SPAppWebUrl;
                context.HttpContext.GetOwinContext().Authentication.Challenge(properties, LoginProvider);
            }
        }
    }
}