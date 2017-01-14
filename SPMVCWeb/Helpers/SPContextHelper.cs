using AspNet.Owin.SharePoint.Addin.Authentication;
using Microsoft.SharePoint.Client;
using SPMVCWeb.Models;
using System;
using System.Collections.Generic;
using System.IdentityModel.Tokens;
using System.Web;
using System.Web.Configuration;

namespace SPMVCWeb.Helpers
{
    internal static class SPContextHelper
    {
        public static RedirectionStatus CheckRedirectionStatus(HttpContextBase httpContext, out Uri redirectUrl)
        {
            return SPContextHelper.CheckRedirectionStatus(httpContext, null, out redirectUrl);
        }

        public static RedirectionStatus CheckRedirectionStatus(HttpContextBase httpContext, Uri requestUrl, out Uri redirectUrl)
        {
            if (httpContext == null)
            {
                throw new ArgumentNullException("httpContext");
            }

            redirectUrl = null;
            bool contextTokenExpired = false;

            Uri spHostUrl = SharePointContext.GetSPHostUrl(httpContext.Request);
            if (spHostUrl == null)
            {
                string spHostUrlString = WebConfigurationManager.AppSettings.Get(SharePointContext.SPHostUrlKey);
                if (!Uri.TryCreate(spHostUrlString, UriKind.Absolute, out spHostUrl))
                {
                    return RedirectionStatus.CanNotRedirect;
                }
            }

            try
            {
                if (SharePointContextProvider.Current.GetSharePointContext(httpContext, spHostUrl) != null)
                {
                    return RedirectionStatus.Ok;
                }
            }
            catch (SecurityTokenExpiredException)
            {
                contextTokenExpired = true;
            }

            const string SPHasRedirectedToSharePointKey = "SPHasRedirectedToSharePoint";

            if (!string.IsNullOrEmpty(httpContext.Request.QueryString[SPHasRedirectedToSharePointKey]) && !contextTokenExpired)
            {
                return RedirectionStatus.CanNotRedirect;
            }

            if (StringComparer.OrdinalIgnoreCase.Equals(httpContext.Request.HttpMethod, "POST"))
            {
                return RedirectionStatus.CanNotRedirect;
            }

            if (requestUrl == null)
            {
                requestUrl = httpContext.Request.Url;
            }
            var queryNameValueCollection = HttpUtility.ParseQueryString(requestUrl.Query);
            // Removes the values that are included in {StandardTokens}, as {StandardTokens} will be inserted at the beginning of the query string.
            queryNameValueCollection.Remove(SharePointContext.SPHostUrlKey);
            queryNameValueCollection.Remove(SharePointContext.SPAppWebUrlKey);
            queryNameValueCollection.Remove(SharePointContext.SPLanguageKey);
            queryNameValueCollection.Remove(SharePointContext.SPClientTagKey);
            queryNameValueCollection.Remove(SharePointContext.SPProductNumberKey);

            // Adds SPHasRedirectedToSharePoint=1.
            queryNameValueCollection.Add(SPHasRedirectedToSharePointKey, "1");

            UriBuilder returnUrlBuilder = new UriBuilder(requestUrl);
            returnUrlBuilder.Query = queryNameValueCollection.ToString();

            // Inserts StandardTokens.
            const string StandardTokens = "{StandardTokens}";
            string returnUrlString = returnUrlBuilder.Uri.AbsoluteUri;
            returnUrlString = returnUrlString.Insert(returnUrlString.IndexOf("?") + 1, StandardTokens + "&");

            // Constructs redirect url.
            string redirectUrlString = TokenHelper.GetAppContextTokenRequestUrl(spHostUrl.GetLeftPart(UriPartial.Path), Uri.EscapeDataString(returnUrlString));
            redirectUrl = new Uri(redirectUrlString, UriKind.Absolute);
            return RedirectionStatus.ShouldRedirect;
        }

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

        public static void ExecuteUserContextQuery<TContext>(ISPContext spContext, Func<TContext, Action> action)
            where TContext : ClientContext
        {
            if (action == null) throw new ArgumentNullException("action");
            if (spContext != null)
            {
                TContext clientContext = (TContext)spContext.CreateUserClientContextForSPHost();
                if (clientContext != null)
                {
                    using (clientContext)
                    {
                        Action result = action.Invoke(clientContext);
                        clientContext.ExecuteQuery();
                        if (result != null)
                        {
                            result.Invoke();
                        }
                    }
                }
            }
        }

        public static void ExecuteAppOnlyClientContextQuery<TContext>(ISPContext spContext, Func<TContext, Action> action)
            where TContext : ClientContext
        {
            if (action == null) throw new ArgumentNullException("action");
            if (spContext != null)
            {
                TContext clientContext = (TContext)spContext.CreateAppOnlyClientContextForSPHost();
                if (clientContext != null)
                {
                    using (clientContext)
                    {
                        Action result = action.Invoke(clientContext);
                        clientContext.ExecuteQuery();
                        if (result != null)
                        {
                            result.Invoke();
                        }
                    }
                }
            }
        }


        public static void ExecuteUserClientContextForSPAppWebQuery<TContext>(ISPContext spContext, Func<TContext, Action> action)
              where TContext : ClientContext
        {
            if (action == null) throw new ArgumentNullException("action");
            if (spContext != null)
            {
                TContext clientContext = (TContext)spContext.CreateUserClientContextForSPAppWeb();
                if (clientContext != null)
                {
                    using (clientContext)
                    {
                        Action result = action.Invoke(clientContext);
                        clientContext.ExecuteQuery();
                        if (result != null)
                        {
                            result.Invoke();
                        }
                    }
                }
            }
        }


        public static void ExecuteAppOnlyClientContextForSPAppWebQuery<TContext>(ISPContext spContext, Func<TContext, Action> action)
            where TContext : ClientContext
        {
            if (action == null) throw new ArgumentNullException("action");
            if (spContext != null)
            {
                TContext clientContext = (TContext)spContext.CreateAppOnlyClientContextForSPAppWeb();
                if (clientContext != null)
                {
                    using (clientContext)
                    {
                        Action result = action.Invoke(clientContext);
                        clientContext.ExecuteQuery();
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
                if (site.IsPropertyAvailable("CompatibilityLevel"))
                    pageContextInfo.LayoutsUrl = GetLayoutsFolder(site.CompatibilityLevel);
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

                pageContextInfo.RegionalInfo = GetSPRegionalInfo(web.RegionalSettings);
            }
            return pageContextInfo;
        }

        public static string GetLayoutsFolder(int compatibilityLevel)
        {
            string layouts = "_layouts";
            if (compatibilityLevel >= 15)
            {
                layouts = string.Concat(layouts, "/", compatibilityLevel);
            }
            return layouts;
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
            if (regionalSettings.TimeZone.IsPropertyAvailable("Information"))
            {
                regionalInfo.TimeZoneBias = regionalSettings.TimeZone.Information.Bias;
            }
            return regionalInfo;
        }
    }
}