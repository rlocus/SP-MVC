using AspNet.Owin.SharePoint.Addin.Authentication;
using Microsoft.SharePoint.Client;
using SPMVCWeb.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web.Configuration;
using System.Web.Mvc;
using Newtonsoft.Json;

namespace SPMVCWeb.Controllers
{
    [SPAuthorize(Permissions = PermissionKind.EmptyMask, SPGroup = "", SiteAdminRequired = false)]
    public class HomeController : Controller
    {
        [SharePointContextFilter]
        public ActionResult Index()
        {
            Site site = null;
            Web web = null;
            InitView((clientContext) =>
            {
                site = clientContext.Site;
                clientContext.Load(site);
                web = clientContext.Web;
                clientContext.Load(web);
            });
            ViewBag.PageContextInfo = getPageContextInfo(site, web);
            return View();
        }

        [SharePointContextFilter]
        public ActionResult About()
        {
            InitView();
            ViewBag.Message = "SP MVC application.";
            return View();
        }

        [SharePointContextFilter]
        public ActionResult Contact()
        {
            InitView();
            ViewBag.Message = "Contact.";
            return View();
        }

        [SharePointContextFilter]
        public ActionResult List(Guid listId, Guid? viewId)
        {
            List list = null;
            View view = null;
            InitView((clientContext) =>
            {
                list = clientContext.Web.Lists.GetById(listId);
                clientContext.Load(list);
                clientContext.Load(list.RootFolder);
                clientContext.Load(list.Fields, fields => fields.Where(f => !f.Hidden && f.Group != "_Hidden"));
                if (viewId == null || default(Guid) == viewId)
                {
                    view = list.DefaultView;
                }
                else
                {
                    view = list.GetViewById(viewId.Value);
                }
                clientContext.Load(view);
                clientContext.Load(view.ViewFields);
            });
            ViewBag.List = new ListInformation(list, view);
            return View();
        }

        private ClientContext GetClientContext()
        {
            var cookieAuthenticationEnabled = string.IsNullOrEmpty(WebConfigurationManager.AppSettings.Get("CookieAuthenticationEnabled")) ? false : Convert.ToBoolean(WebConfigurationManager.AppSettings.Get("CookieAuthenticationEnabled"));
            if (cookieAuthenticationEnabled)
            {
                var spContext = SPContextProvider.Get(User as System.Security.Claims.ClaimsPrincipal);
                if (spContext != null)
                {
                    return spContext.CreateUserClientContextForSPHost();
                }
            }
            else
            {
                var spContext = SharePointContextProvider.Current.GetSharePointContext(HttpContext);
                if (spContext != null)
                {
                    return spContext.CreateUserClientContextForSPHost();
                }
            }
            return null;
        }

        private void InitView(Action<ClientContext> action = null)
        {
            User spUser = null;
            ClientContext clientContext = GetClientContext();
            if (clientContext != null)
            {
                using (clientContext)
                {
                    spUser = clientContext.Web.CurrentUser;
                    clientContext.Load(spUser);
                    if (action != null)
                    {
                        action.Invoke(clientContext);
                    }
                    clientContext.ExecuteQuery();
                    ViewBag.User = new UserInformation(spUser);
                    ViewBag.FormDigest = clientContext.GetFormDigestDirect().DigestValue;
                }
            }
        }

        private SPPageContextInfo getPageContextInfo(Site site, Web web)
        {
            SPPageContextInfo pageContextInfo = new SPPageContextInfo();
            if (site != null)
            {
                if (site.IsPropertyAvailable("SiteServerRelativeUrl"))
                    pageContextInfo.SiteServerRelativeUrl = site.ServerRelativeUrl;
                if (site.IsPropertyAvailable("Url"))
                    pageContextInfo.SiteAbsoluteUrl = site.Url;
            }
            if (web != null)
            {
                if (web.IsPropertyAvailable("ServerRelativeUrl"))
                    pageContextInfo.WebServerRelativeUrl = web.ServerRelativeUrl;
                if (web.IsPropertyAvailable("Url"))
                    pageContextInfo.WebAbsoluteUrl = web.Url;
                if (web.IsPropertyAvailable("Language"))
                    pageContextInfo.WebLanguage = web.Language;
                if (web.IsPropertyAvailable("SiteLogoUrl"))
                    pageContextInfo.WebLogoUrl = web.SiteLogoUrl;
                if (web.IsPropertyAvailable("EffectiveBasePermissions"))
                {
                    var permissions = new List<int>();
                    foreach (var pk in (PermissionKind[])Enum.GetValues(typeof(PermissionKind)))
                    {
                        if (web.EffectiveBasePermissions.Has(pk) && pk != PermissionKind.EmptyMask)
                        {
                            permissions.Add((int)pk);
                        }
                    }
                    pageContextInfo.WebPermMasks = JsonConvert.SerializeObject(permissions);
                }
                if (web.IsPropertyAvailable("Title"))
                    pageContextInfo.WebTitle = web.Title;
                if (web.IsPropertyAvailable("UIVersion"))
                    pageContextInfo.WebUIVersion = web.UIVersion;

                User user = web.CurrentUser;
                if (user.IsPropertyAvailable("Id"))
                    pageContextInfo.UserId = user.Id;
                if (user.IsPropertyAvailable("LoginName"))
                    pageContextInfo.UserLoginName = user.LoginName;
            }
            return pageContextInfo;
        }
    }
}
