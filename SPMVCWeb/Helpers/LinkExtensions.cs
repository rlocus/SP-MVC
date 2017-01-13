using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.Web.Mvc.Html;
using System.Web.Routing;
using AspNet.Owin.SharePoint.Addin.Authentication;

namespace SPMVCWeb.Helpers
{
    public static class LinkExtensions
    {
        public static MvcHtmlString SPActionLink(this HtmlHelper htmlHelper, string linkText, string actionName, string controllerName, RouteValueDictionary routeValues, IDictionary<string, object> htmlAttributes)
        {
            var spContext = SPContextHelper.GetSPContext(htmlHelper.ViewContext.HttpContext);
            if (spContext != null)
            {
                if (routeValues == null)
                {
                    routeValues = new RouteValueDictionary();
                }
                if (spContext.SPHostUrl != null)
                {
                    if (routeValues.ContainsKey(SharePointContext.SPHostUrlKey))
                    {
                        routeValues[SharePointContext.SPHostUrlKey] = spContext.SPHostUrl.GetLeftPart(UriPartial.Path);
                    }
                    else
                    {
                        routeValues.Add(SharePointContext.SPHostUrlKey, spContext.SPHostUrl.GetLeftPart(UriPartial.Path));
                    }
                }
                if (spContext.SPAppWebUrl != null)
                {
                    if (routeValues.ContainsKey(SharePointContext.SPAppWebUrlKey))
                    {
                        routeValues[SharePointContext.SPAppWebUrlKey] = spContext.SPAppWebUrl.GetLeftPart(UriPartial.Path);
                    }
                    else
                    {
                        routeValues.Add(SharePointContext.SPAppWebUrlKey, spContext.SPAppWebUrl.GetLeftPart(UriPartial.Path));
                    }
                }
            }
            return htmlHelper.ActionLink(linkText, actionName, controllerName, routeValues, htmlAttributes);
        }
    }
}