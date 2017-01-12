﻿using AspNet.Owin.SharePoint.Addin.Authentication;
using AspNet.Owin.SharePoint.Addin.Authentication.Middleware;
using Microsoft.Owin.Security;
using System;
using System.Configuration;
using System.Net.Http;
using System.Security.Claims;
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
            if (string.IsNullOrEmpty(spHostUrlString))
            {
                spHostUrlString = ConfigurationManager.AppSettings["SPHostUrl"];
            }
            Uri spHostUrl;
            if (!Uri.TryCreate(spHostUrlString, UriKind.Absolute, out spHostUrl) &&
                                            (spHostUrl.Scheme == Uri.UriSchemeHttp || spHostUrl.Scheme == Uri.UriSchemeHttps))
            {
                throw new Exception(string.Format("Unable to determine {0}.", SharePointContext.SPHostUrlKey));
            }
            return new ChallengeResult(SPAddinAuthenticationDefaults.AuthenticationType, spHostUrl.ToString(), returnUrl);
        }

        [HttpGet]
        public ActionResult Logout()
        {
            if (Request.IsAuthenticated)
            {
                var spContext = SPContextProvider.Get(HttpContext.User as ClaimsPrincipal);
                HttpContext.GetOwinContext()
                    .Authentication.SignOut(new AuthenticationProperties() { RedirectUri = "/" }, SPAddinAuthenticationDefaults.AuthenticationType);
                if (spContext.SPAppWebUrl != null)
                    return new RedirectResult(string.Format("{0}/_layouts/closeConnection.aspx?loginasanotheruser=true", spContext.SPAppWebUrl.GetLeftPart(UriPartial.Path).TrimEnd('/')));
            }
            return new RedirectResult(null);
        }

        public void EndSession()
        {
            HttpContext.GetOwinContext().Authentication.SignOut(SPAddinAuthenticationDefaults.AuthenticationType);
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
                //properties.Dictionary[SharePointContext.SPAppWebUrlKey] = SPAppWebUrl;
                context.HttpContext.GetOwinContext().Authentication.Challenge(properties, LoginProvider);
            }
        }
    }
}