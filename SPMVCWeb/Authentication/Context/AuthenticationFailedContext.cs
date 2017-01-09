using AspNet.Owin.SharePoint.Addin.Authentication.Middleware;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using Microsoft.Owin;

namespace AspNet.Owin.SharePoint.Addin.Authentication.Context
{
    public class AuthenticationFailedContext : BaseSharePointAuthenticationContext
    {
        public AuthenticationFailedContext(IOwinContext context, SPAddInAuthenticationOptions options)
               : base(context, options)
        {
        }
    }
}