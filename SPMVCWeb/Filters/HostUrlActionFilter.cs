using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using AspNet.Owin.SharePoint.Addin.Authentication;

namespace SPMVCWeb.Filters
{
    public class HostUrlActionFilter : ActionFilterAttribute
    {
        public override void OnResultExecuting(ResultExecutingContext filterContext)
        {
            filterContext.Controller.ViewBag.SPHostUrl = SharePointContext.GetSPHostUrl(filterContext.HttpContext.Request);
            filterContext.Controller.ViewBag.SPAppWebUrl = filterContext.HttpContext.Request[SharePointContext.SPAppWebUrlKey];
        }
    }
}