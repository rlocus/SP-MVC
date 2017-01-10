using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Configuration;
using System.Web.Mvc;

namespace SPMVCWeb.Controllers
{
    public class SPAuthorizeAttribute : AuthorizeAttribute
    {
        // Custom property
        public string SPGroup { get; set; }

        public string SPRole { get; set; }

        protected override bool AuthorizeCore(HttpContextBase httpContext)
        {
            var cookieAuthenticationEnabled = string.IsNullOrEmpty(WebConfigurationManager.AppSettings.Get("CookieAuthenticationEnabled")) ? false : Convert.ToBoolean(WebConfigurationManager.AppSettings.Get("CookieAuthenticationEnabled"));
            bool authorized = !cookieAuthenticationEnabled || base.AuthorizeCore(httpContext);
            return authorized;
        }

        //protected override void HandleUnauthorizedRequest(AuthorizationContext filterContext)
        //{
        //    var cookieAuthenticationEnabled = string.IsNullOrEmpty(WebConfigurationManager.AppSettings.Get("CookieAuthenticationEnabled")) ? false : Convert.ToBoolean(WebConfigurationManager.AppSettings.Get("CookieAuthenticationEnabled"));

        //    base.HandleUnauthorizedRequest(filterContext);
        //}
    }
}