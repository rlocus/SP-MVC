using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using System.Web;
using System.Web.Mvc;

namespace SPMVCWeb.Controllers
{
    public class PersonalData
    {
        public string Initials { get; set; }
        public string Name { get; set; }
        public string Title { get; set; }
    }

    public class HomeController : Controller
    {
        [SharePointContextFilter]
        public ActionResult Index()
        {
            InitView();
            return View();
        }

        public ActionResult About()
        {
            InitView();
            ViewBag.Message = "SP MVC application.";
            return View();
        }

        public ActionResult Contact()
        {
            InitView();
            ViewBag.Message = "Contact.";
            return View();
        }

        private ClientContext GetClientContext()
        {
            var spContext = SharePointContextProvider.Current.GetSharePointContext(HttpContext);
            if (spContext != null)
            {
                return spContext.CreateUserClientContextForSPHost();
            }
            return null;
        }

        private void InitView()
        {
            User spUser = null;
            ClientContext clientContext = GetClientContext();
            if (clientContext != null)
                using (clientContext)
                {
                    if (clientContext != null)
                    {
                        spUser = clientContext.Web.CurrentUser;
                        clientContext.Load(spUser, user => user.Title);
                        clientContext.ExecuteQuery();
                        ViewBag.User = new PersonalData
                        {
                            Initials = new Regex(@"(\b[a-zA-Z])[a-zA-Z]* ?").Replace(spUser.Title, "$1"),
                            Name = spUser.Title
                        };
                        ViewBag.FormDigest = clientContext.GetFormDigestDirect().DigestValue;
                    }
                }
        }
    }
}
