﻿using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace SPMVCWeb.Controllers
{
    public class HomeController : Controller
    {
        [SharePointContextFilter]
        public ActionResult Index()
        {
            User spUser = null;
            var spContext = SharePointContextProvider.Current.GetSharePointContext(HttpContext);
            using (var clientContext = spContext.CreateUserClientContextForSPHost())
            {
                if (clientContext != null)
                {
                    spUser = clientContext.Web.CurrentUser;
                    clientContext.Load(spUser, user => user.Title);
                    clientContext.ExecuteQuery();
                    ViewBag.UserName = spUser.Title;
                }
            }
            return View();
        }

        public ActionResult About()
        {
            ViewBag.Message = "SP MVC application.";
            return View();
        }

        public ActionResult Contact()
        {
            ViewBag.Message = "Contact.";
            return View();
        }
    }
}
