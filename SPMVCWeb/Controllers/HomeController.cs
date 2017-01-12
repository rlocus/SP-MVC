using Microsoft.SharePoint.Client;
using SPMVCWeb.Helpers;
using SPMVCWeb.Models;
using System;
using System.Linq;
using System.Web.Mvc;

namespace SPMVCWeb.Controllers
{
    [SPAuthorize(Permissions = PermissionKind.EmptyMask, SPGroup = "", SiteAdminRequired = false)]
    public class HomeController : Controller
    {
        [SharePointContextFilter]
        public ActionResult Index()
        {
            var spContext = SPContextHelper.GetSPContext(this.HttpContext);
            if (spContext != null)
            {
                SPContextHelper.RunWithContext(spContext, (clientContext) =>
                {
                    User spUser = clientContext.Web.CurrentUser;
                    clientContext.Load(spUser);
                    Site site = clientContext.Site;
                    clientContext.Load(site);
                    Web web = clientContext.Web;
                    clientContext.Load(web);
                    clientContext.Load(web.RegionalSettings);
                    clientContext.Load(web.RegionalSettings.TimeZone);
                    return () =>
                    {
                        ViewBag.User = new UserInformation(spUser);
                        ViewBag.FormDigest = clientContext.GetFormDigestDirect().DigestValue;
                        SPPageContextInfo pageContextInfo =  SPContextHelper.GetPageContextInfo(site, web);
                        if (spContext.SPAppWebUrl != null)
                        {
                            pageContextInfo.AppWebUrl = spContext.SPAppWebUrl.AbsoluteUri;
                        }
                        ViewBag.PageContextInfo = pageContextInfo;
                    };
                });
            }
            return View();
        }

        [SharePointContextFilter]
        public ActionResult About()
        {
            var spContext = SPContextHelper.GetSPContext(this.HttpContext);
            if (spContext != null)
            {
                SPContextHelper.RunWithContext(spContext, (clientContext) =>
                {
                    User spUser = clientContext.Web.CurrentUser;
                    clientContext.Load(spUser);
                    Site site = clientContext.Site;
                    clientContext.Load(site);
                    Web web = clientContext.Web;
                    clientContext.Load(web);
                    return () =>
                    {
                        ViewBag.User = new UserInformation(spUser);
                        SPPageContextInfo pageContextInfo = SPContextHelper.GetPageContextInfo(site, web);
                        if (spContext.SPAppWebUrl != null)
                        {
                            pageContextInfo.AppWebUrl = spContext.SPAppWebUrl.AbsoluteUri;
                        }
                        ViewBag.PageContextInfo = pageContextInfo;
                    };
                });
            }
            ViewBag.Message = "SP MVC application.";
            return View();
        }

        [SharePointContextFilter]
        public ActionResult Contact()
        {
            var spContext = SPContextHelper.GetSPContext(this.HttpContext);
            if (spContext != null)
            {
                SPContextHelper.RunWithContext(spContext, (clientContext) =>
                {
                    User spUser = clientContext.Web.CurrentUser;
                    clientContext.Load(spUser);
                    Site site = clientContext.Site;
                    clientContext.Load(site);
                    Web web = clientContext.Web;
                    clientContext.Load(web);
                    return () =>
                    {
                        ViewBag.User = new UserInformation(spUser);
                        SPPageContextInfo pageContextInfo = SPContextHelper.GetPageContextInfo(site, web);
                        if (spContext.SPAppWebUrl != null)
                        {
                            pageContextInfo.AppWebUrl = spContext.SPAppWebUrl.AbsoluteUri;
                        }
                        ViewBag.PageContextInfo = pageContextInfo;
                    };
                });
            }
            ViewBag.Message = "Contact.";
            return View();
        }

        [SharePointContextFilter]
        public ActionResult List(Guid listId, Guid? viewId)
        {
            var spContext = SPContextHelper.GetSPContext(this.HttpContext);
            if (spContext != null)
            {
                SPContextHelper.RunWithContext(spContext, (clientContext) =>
                {
                    User spUser = clientContext.Web.CurrentUser;
                    clientContext.Load(spUser);
                    Site site = clientContext.Site;
                    clientContext.Load(site);
                    Web web = clientContext.Web;
                    clientContext.Load(web);
                    clientContext.Load(web.RegionalSettings);
                    clientContext.Load(web.RegionalSettings.TimeZone);

                    List list = clientContext.Web.Lists.GetById(listId);
                    View view;
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

                    return () =>
                    {
                        ViewBag.User = new UserInformation(spUser);
                        ViewBag.FormDigest = clientContext.GetFormDigestDirect().DigestValue;
                        SPPageContextInfo pageContextInfo = SPContextHelper.GetPageContextInfo(site, web);
                        if (spContext.SPAppWebUrl != null)
                        {
                            pageContextInfo.AppWebUrl = spContext.SPAppWebUrl.AbsoluteUri;
                        }
                        ViewBag.PageContextInfo = pageContextInfo; ViewBag.List = new ListInformation(list, view);
                    };
                });
            }
            return View();
        }
    }
}
