using AspNet.Owin.SharePoint.Addin.Authentication.Context;
using System;
using System.Net;
using System.Net.Http;
using System.Web;
using ActionFilterAttribute = System.Web.Http.Filters.ActionFilterAttribute;

namespace SPMVCWeb.Filters
{
    public class SharePointContextWebAPIFilterAttribute : ActionFilterAttribute
    {
        public override void OnActionExecuting(System.Web.Http.Controllers.HttpActionContext actionContext)
        {
            if (actionContext == null)
            {
                throw new ArgumentNullException("actionContext");
            }

            Uri redirectUrl;
            switch (SharePointContextProvider.CheckRedirectionStatus(HttpContext.Current, out redirectUrl))
            {
                case RedirectionStatus.Ok:
                    return;
                case RedirectionStatus.ShouldRedirect:
                    var response = actionContext.Request.CreateResponse(HttpStatusCode.Redirect);
                    response.Headers.Add("Location", redirectUrl.AbsoluteUri);
                    actionContext.Response = response;
                    break;
                case RedirectionStatus.CanNotRedirect:
                    actionContext.Response = actionContext.Request.CreateErrorResponse(HttpStatusCode.MethodNotAllowed, "Access denied.");
                    break;
            }
        }
    }
}