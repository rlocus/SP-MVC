﻿using Microsoft.AspNet.SignalR;
using Owin;
using Microsoft.Owin;
using Microsoft.Owin.Security;
using Microsoft.Owin.Security.Cookies;
using System.Configuration;
using AspNet.Owin.SharePoint.Addin.Authentication.Middleware;
using AspNet.Owin.SharePoint.Addin.Authentication.Provider;
using System.Threading.Tasks;
using System.Web.Configuration;
using System;

namespace SPMVCWeb
{
    public class Startup : Hub
    {
        public void Configuration(IAppBuilder app)
        {
            var cookieAuthenticationEnabled = !string.IsNullOrEmpty(WebConfigurationManager.AppSettings.Get("CookieAuthenticationEnabled")) && Convert.ToBoolean(WebConfigurationManager.AppSettings.Get("CookieAuthenticationEnabled"));
            if (cookieAuthenticationEnabled)
            {
                var cookieAuth = new CookieAuthenticationOptions
                {
                    AuthenticationType = SPAddinAuthenticationDefaults.AuthenticationType,
                    LoginPath = new PathString("/login"),
                    LogoutPath = new PathString("/logout"),
                    Provider = new AddInCookieAuthenticationProvider()
                };
                app.SetDefaultSignInAsAuthenticationType(cookieAuth.AuthenticationType);
                app.UseCookieAuthentication(cookieAuth);
                string clientId = ConfigurationManager.AppSettings["ClientId"];
                app.UseSPAddinAuthentication(new SPAddInAuthenticationOptions
                {
                    ClientId = string.IsNullOrEmpty(clientId) ? Guid.Empty : new Guid(clientId),
                    SPHostUrl = ConfigurationManager.AppSettings["SPHostUrl"],
                    Provider = new SPAddinAuthenticationProvider()
                    {
                        OnAuthenticated = (context) => Task.FromResult<object>(null)
                    }
                });
            }
        }
    }
}