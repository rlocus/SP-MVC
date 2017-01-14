using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.Web.Routing;

namespace SPMVCWeb
{
    public class RouteConfig
    {
        public static void RegisterRoutes(RouteCollection routes)
        {
            routes.IgnoreRoute("{resource}.axd/{*pathInfo}");

            routes.MapRoute(
                name: "Log in",
                url: "login",
                defaults: new { controller = "Account", action= "Login" }
            );

            routes.MapRoute(
               name: "Log out",
               url: "logout",
               defaults: new { controller = "Account", action = "Logout" }
           );

            routes.MapRoute(
                name: "Default",
                url: "{controller}/{action}/{id}",
                defaults: new { controller = "Home", action = "Index", id = UrlParameter.Optional }
            );

            routes.MapRoute(
                name: "DefaultApi",
                url: "api/{controller}/{id}",
                defaults: new { id = UrlParameter.Optional }
                ).RouteHandler = new SessionRouteHandler();
        }
    }
}
