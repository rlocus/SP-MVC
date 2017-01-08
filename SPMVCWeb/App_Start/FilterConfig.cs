using SPMVCWeb.Filters;
using System.Web;
using System.Web.Mvc;

namespace SPMVCWeb
{
    public class FilterConfig
    {
        public static void RegisterGlobalFilters(GlobalFilterCollection filters)
        {
            //filters.Add(new AuthorizeAttribute());
            //filters.Add(new HostUrlActionFilter());
            filters.Add(new HandleErrorAttribute());           
        }
    }
}
