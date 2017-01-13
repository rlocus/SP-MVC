using AspNet.Owin.SharePoint.Addin.Authentication;
using SPMVCWeb.Helpers;
using System;
using System.Web.Configuration;
using System.Web.Mvc;

namespace SPMVCWeb
{
    /// <summary>
    /// SharePoint action filter attribute.
    /// </summary>
    public class SharePointContextFilterAttribute : ActionFilterAttribute
    {
        public override void OnActionExecuting(ActionExecutingContext filterContext)
        {
            if (filterContext == null)
            {
                throw new ArgumentNullException("filterContext");
            }

            var cookieAuthenticationEnabled = string.IsNullOrEmpty(WebConfigurationManager.AppSettings.Get("CookieAuthenticationEnabled")) ? false : Convert.ToBoolean(WebConfigurationManager.AppSettings.Get("CookieAuthenticationEnabled"));
            if (cookieAuthenticationEnabled) return;

            Uri redirectUrl;
            switch (SPContextHelper.CheckRedirectionStatus(filterContext.HttpContext, out redirectUrl))
            {
                case RedirectionStatus.Ok:
                    return;
                case RedirectionStatus.ShouldRedirect:
                    filterContext.Result = new RedirectResult(redirectUrl.AbsoluteUri);
                    break;
                case RedirectionStatus.CanNotRedirect:
                    filterContext.Result = new ViewResult { ViewName = "Error" };
                    break;
            }
        }
    }
}
