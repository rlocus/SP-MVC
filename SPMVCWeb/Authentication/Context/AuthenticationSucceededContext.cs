using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using Microsoft.Owin;
using AspNet.Owin.SharePoint.Addin.Authentication.Middleware;

namespace AspNet.Owin.SharePoint.Addin.Authentication.Context
{
    public class AuthenticationSucceededContext : BaseSharePointAuthenticationContext
    {
        public AuthenticationSucceededContext(IOwinContext context, SPAddInAuthenticationOptions options)
               : base(context, options)
        {
        }

        public SharePointContext SharePointContext { get; set; }
    }
}