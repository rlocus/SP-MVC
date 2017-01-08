using AspNet.Owin.SharePoint.Addin.Authentication.Context;
using Microsoft.SharePoint.Client;
using SPMVCWeb.Models;
using System;
using System.Linq;
using System.Web.Mvc;

namespace SPMVCWeb.Controllers
{
    //[Authorize]
    public class HomeController : Controller
    {
        [SharePointContextFilter]
        public ActionResult Index()
        {
            InitView();
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
            //var spContext = SPContextProvider.Get(User as ClaimsPrincipal);
            var spContext = SharePointContextProvider.Current.GetSharePointContext(HttpContext);
            if (spContext != null)
            {
                return spContext.CreateUserClientContextForSPHost();
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
    }
}
