using AspNet.Owin.SharePoint.Addin.Authentication;
using Microsoft.SharePoint.Client;
using SPMVCWeb.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Configuration;

namespace SPMVCWeb.Helpers
{
    static class SPContextHelper
    {
        public static ISPContext GetSPContext(HttpContextBase httpContext)
        {
            var cookieAuthenticationEnabled = string.IsNullOrEmpty(WebConfigurationManager.AppSettings.Get("CookieAuthenticationEnabled")) ? false : Convert.ToBoolean(WebConfigurationManager.AppSettings.Get("CookieAuthenticationEnabled"));
            if (cookieAuthenticationEnabled)
            {
                return SPContextProvider.Get(httpContext.User as System.Security.Claims.ClaimsPrincipal);
            }
            else
            {
                Uri spHostUrl = SharePointContext.GetSPHostUrl(httpContext.Request);
                if (spHostUrl == null)
                {
                    string spHostUrlString = WebConfigurationManager.AppSettings.Get(SharePointContext.SPHostUrlKey);
                    if (!Uri.TryCreate(spHostUrlString, UriKind.Absolute, out spHostUrl))
                    {
                      
                    }
                }
                return SharePointContextProvider.Current.GetSharePointContext(httpContext, spHostUrl);
            }
        }

        public static void RunWithContext(ISPContext spContext, Func<ClientContext, Action> action)
        {
            if (action == null) throw new ArgumentNullException("action");
            if (spContext != null)
            {
                ClientContext clientContext = spContext.CreateUserClientContextForSPHost();
                if (clientContext != null)
                {
                    using (clientContext)
                    {
                        //User spUser = clientContext.Web.CurrentUser;
                        //clientContext.Load(spUser);
                        Action result = action.Invoke(clientContext);
                        clientContext.ExecuteQuery();
                        //ViewBag.User = new UserInformation(spUser);
                        //ViewBag.FormDigest = clientContext.GetFormDigestDirect().DigestValue;
                        if (result != null)
                        {
                            result.Invoke();
                        }
                    }
                }
            }
        }

        public static SPPageContextInfo GetPageContextInfo(Site site, Web web)
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
                    pageContextInfo.WebPermMasks = permissions.ToArray();
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

                if (web.IsPropertyAvailable("RegionalSettings"))
                {
                    pageContextInfo.RegionalInfo = GetSPRegionalInfo(web.RegionalSettings);
                }
            }
            return pageContextInfo;
        }

        public static SPRegionalInfo GetSPRegionalInfo(RegionalSettings regionalSettings)
        {
            SPRegionalInfo regionalInfo = new SPRegionalInfo();
            if (regionalSettings.IsPropertyAvailable("AM"))
            {
                regionalInfo.AM = regionalSettings.AM;
            }
            if (regionalSettings.IsPropertyAvailable("AdjustHijriDays"))
            {
                regionalInfo.AdjustHijriDays = regionalSettings.AdjustHijriDays;
            }
            if (regionalSettings.IsPropertyAvailable("AlternateCalendarType"))
            {
                regionalInfo.AlternateCalendarType = regionalSettings.AlternateCalendarType;
            }
            if (regionalSettings.IsPropertyAvailable("CalendarType"))
            {
                regionalInfo.CalendarType = regionalSettings.CalendarType;
            }
            if (regionalSettings.IsPropertyAvailable("Collation"))
            {
                regionalInfo.Collation = regionalSettings.Collation;
            }
            if (regionalSettings.IsPropertyAvailable("CollationLCID"))
            {
                regionalInfo.CollationLCID = regionalSettings.CollationLCID;
            }
            if (regionalSettings.IsPropertyAvailable("DateFormat"))
            {
                regionalInfo.DateFormat = regionalSettings.DateFormat;
            }
            if (regionalSettings.IsPropertyAvailable("DateSeparator"))
            {
                regionalInfo.DateSeparator = regionalSettings.DateSeparator;
            }
            if (regionalSettings.IsPropertyAvailable("DecimalSeparator"))
            {
                regionalInfo.DecimalSeparator = regionalSettings.DecimalSeparator;
            }
            if (regionalSettings.IsPropertyAvailable("DigitGrouping"))
            {
                regionalInfo.DigitGrouping = regionalSettings.DigitGrouping;
            }
            if (regionalSettings.IsPropertyAvailable("FirstDayOfWeek"))
            {
                regionalInfo.FirstDayOfWeek = regionalSettings.FirstDayOfWeek;
            }
            if (regionalSettings.IsPropertyAvailable("FirstWeekOfYear"))
            {
                regionalInfo.FirstWeekOfYear = regionalSettings.FirstWeekOfYear;
            }
            if (regionalSettings.IsPropertyAvailable("IsEastAsia"))
            {
                regionalInfo.IsEastAsia = regionalSettings.IsEastAsia;
            }
            if (regionalSettings.IsPropertyAvailable("IsRightToLeft"))
            {
                regionalInfo.IsRightToLeft = regionalSettings.IsRightToLeft;
            }
            if (regionalSettings.IsPropertyAvailable("IsUIRightToLeft"))
            {
                regionalInfo.IsUIRightToLeft = regionalSettings.IsUIRightToLeft;
            }
            if (regionalSettings.IsPropertyAvailable("ListSeparator"))
            {
                regionalInfo.ListSeparator = regionalSettings.ListSeparator;
            }
            if (regionalSettings.IsPropertyAvailable("LocaleId"))
            {
                regionalInfo.LocaleId = regionalSettings.LocaleId;
            }
            if (regionalSettings.IsPropertyAvailable("NegNumberMode"))
            {
                regionalInfo.NegNumberMode = regionalSettings.NegNumberMode;
            }
            if (regionalSettings.IsPropertyAvailable("NegativeSign"))
            {
                regionalInfo.NegativeSign = regionalSettings.NegativeSign;
            }
            if (regionalSettings.IsPropertyAvailable("PM"))
            {
                regionalInfo.PM = regionalSettings.PM;
            }
            if (regionalSettings.IsPropertyAvailable("PositiveSign"))
            {
                regionalInfo.PositiveSign = regionalSettings.PositiveSign;
            }
            if (regionalSettings.IsPropertyAvailable("ShowWeeks"))
            {
                regionalInfo.ShowWeeks = regionalSettings.ShowWeeks;
            }
            if (regionalSettings.IsPropertyAvailable("ThousandSeparator"))
            {
                regionalInfo.ThousandSeparator = regionalSettings.ThousandSeparator;
            }
            if (regionalSettings.IsPropertyAvailable("Time24"))
            {
                regionalInfo.Time24 = regionalSettings.Time24;
            }
            if (regionalSettings.IsPropertyAvailable("TimeMarkerPosition"))
            {
                regionalInfo.TimeMarkerPosition = regionalSettings.TimeMarkerPosition;
            }
            if (regionalSettings.IsPropertyAvailable("TimeSeparator"))
            {
                regionalInfo.TimeSeparator = regionalSettings.TimeSeparator;
            }
            if (regionalSettings.IsPropertyAvailable("WorkDayEndHour"))
            {
                regionalInfo.WorkDayEndHour = regionalSettings.WorkDayEndHour;
            }
            if (regionalSettings.IsPropertyAvailable("WorkDayStartHour"))
            {
                regionalInfo.WorkDayStartHour = regionalSettings.WorkDayStartHour;
            }
            if (regionalSettings.IsPropertyAvailable("WorkDays"))
            {
                regionalInfo.WorkDays = regionalSettings.WorkDays;
            }
            return regionalInfo;
        }
    }
}