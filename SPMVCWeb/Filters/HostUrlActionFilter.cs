using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace SPMVCWeb.Filters
{
    public class HostUrlActionFilter : ActionFilterAttribute
    {
        public override void OnResultExecuting(ResultExecutingContext filterContext)
        {
            filterContext.Controller.ViewBag.HostUrl = filterContext.HttpContext.Request.QueryString["h"];
        }
    }
}