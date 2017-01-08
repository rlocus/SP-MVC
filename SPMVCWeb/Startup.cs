using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using Microsoft.AspNet.SignalR;
using Owin;
using Microsoft.Owin;
using Microsoft.Owin.Security;
using Microsoft.Owin.Security.Cookies;
using System.Configuration;
using AspNet.Owin.SharePoint.Addin.Authentication.Middleware;
using AspNet.Owin.SharePoint.Addin.Authentication.Context;
using AspNet.Owin.SharePoint.Addin.Authentication.Provider;

namespace SPMVCWeb
{
    public class Startup : Hub
    {
        public void Configuration(IAppBuilder app)
        {
            //var cookieAuth = new CookieAuthenticationOptions
            //{
            //    LoginPath = new PathString("/Auth/Login"),
            //    Provider = new AddInCookieAuthenticationProvider()
            //};

            //app.SetDefaultSignInAsAuthenticationType(cookieAuth.AuthenticationType);
            //app.UseCookieAuthentication(cookieAuth);

            //app.UseSPAddinAuthentication(new SPAddInAuthenticationOptions
            //{
            //    ClientId = ConfigurationManager.AppSettings["ClientId"]
            //});
        }
    }
}