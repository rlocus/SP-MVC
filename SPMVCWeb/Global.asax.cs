using SPMVCWeb.Controllers;
using System;
using System.Net;
using System.Web;
using System.Web.Mvc;
using System.Web.Optimization;
using System.Web.Routing;

namespace SPMVCWeb
{
    public class MvcApplication : HttpApplication
    {
        protected void Application_Start()
        {
            AreaRegistration.RegisterAllAreas();
            FilterConfig.RegisterGlobalFilters(GlobalFilters.Filters);
            RouteConfig.RegisterRoutes(RouteTable.Routes);
            BundleConfig.RegisterBundles(BundleTable.Bundles);
        }

        protected void Application_Error()
        {
            var exception = Server.GetLastError();
            Server.ClearError();
            if (exception is UnauthorizedAccessException || exception is Microsoft.SharePoint.Client.ServerUnauthorizedAccessException || (exception is WebException && ((WebException)exception).Status == WebExceptionStatus.ProtocolError))
            {
                var routeData = new RouteData();
                routeData.Values.Add("controller", "Account");
                routeData.Values.Add("action", "UnauthorizedAccess");
                routeData.Values.Add("exception", exception);

                Response.TrySkipIisCustomErrors = true;
                IController controller = new AccountController();
                controller.Execute(new RequestContext(new HttpContextWrapper(Context), routeData));
            }
        }
    }
}
