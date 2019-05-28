using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Claims;
using System.ServiceModel.Security;
using System.Web;
using System.Web.Configuration;
using System.Web.Mvc;
using AspNet.Owin.SharePoint.Addin.Authentication;
using Microsoft.SharePoint.Client;
using AuthorizationContext = System.Web.Mvc.AuthorizationContext;

namespace SPMVCWeb.Controllers
{
    public class SPAuthorizeAttribute : AuthorizeAttribute
    {
        // Custom property
        public string SPGroup { get; set; }

        public PermissionKind Permissions { get; set; }

        public bool SiteAdminRequired { get; set; }

        public SPAuthorizeAttribute()
        {
            Permissions = PermissionKind.EmptyMask;
        }

        protected override bool AuthorizeCore(HttpContextBase httpContext)
        {
            var cookieAuthenticationEnabled = !string.IsNullOrEmpty(WebConfigurationManager.AppSettings.Get("CookieAuthenticationEnabled")) && Convert.ToBoolean(WebConfigurationManager.AppSettings.Get("CookieAuthenticationEnabled"));
            bool authorized = !cookieAuthenticationEnabled || base.AuthorizeCore(httpContext);
            if (authorized)
            {
                if (!string.IsNullOrEmpty(SPGroup) || Permissions != PermissionKind.EmptyMask || SiteAdminRequired)
                {
                    ClientContext clientContext = null;
                    if (cookieAuthenticationEnabled && httpContext.User.Identity.IsAuthenticated)
                    {
                        var spContext = SPContextProvider.Get(httpContext.User.Identity as ClaimsIdentity);
                        if (spContext != null)
                        {
                            clientContext = spContext.CreateUserClientContextForSPHost();
                        }
                    }
                    else
                    {
                        var spContext = SharePointContextProvider.Current.GetSharePointContext();
                        if (spContext != null)
                        {
                            clientContext = spContext.CreateUserClientContextForSPHost();
                        }
                    }
                    if (clientContext != null)
                    {
                        User user = clientContext.Web.CurrentUser;
                        ClientResult<bool> hasPermissions;
                        List<Func<bool>> checkers = new List<Func<bool>>();
                        if (SiteAdminRequired)
                        {
                            clientContext.Load(user, u => u.IsSiteAdmin);
                            checkers.Add(() => user.IsSiteAdmin);
                        }
                        if (!string.IsNullOrEmpty(SPGroup))
                        {
                            var groups = clientContext.LoadQuery(user.Groups.Include(g => g.LoginName));
                            checkers.Add(() =>
                            {
                                return groups.Any(g => g.LoginName == SPGroup);
                            });
                        }
                        if (Permissions != PermissionKind.EmptyMask)
                        {
                            var perm = new BasePermissions();
                            perm.Set(Permissions);
                            hasPermissions = clientContext.Web.DoesUserHavePermissions(perm);
                            checkers.Add(() => hasPermissions.Value);
                        }
                        if (checkers.Count > 0)
                        {
                            clientContext.ExecuteQuery();
                            authorized = checkers.All(c => c());
                            if (!authorized)
                            {
                                throw new UnauthorizedAccessException();
                            }
                        }
                    }
                }
            }
            return authorized;
        }

        //protected override void HandleUnauthorizedRequest(AuthorizationContext filterContext)
        //{
        //    base.HandleUnauthorizedRequest(filterContext);
        //}
    }
}