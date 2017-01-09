using AspNet.Owin.SharePoint.Addin.Authentication.Middleware;
using Microsoft.Owin.Security;
using Microsoft.Owin.Security.Provider;
using System;
using System.Web;
using Microsoft.Owin;

namespace AspNet.Owin.SharePoint.Addin.Authentication.Context
{
    public class BaseSharePointAuthenticationContext : BaseContext
    {
        public BaseSharePointAuthenticationContext(IOwinContext context, SPAddInAuthenticationOptions options)
            : base(context)
        {
            if (options == null)
            {
                throw new ArgumentNullException(nameof(options));
            }

            Options = options;
        }

        public SPAddInAuthenticationOptions Options { get; }
    }
}