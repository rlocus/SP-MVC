using AspNet.Owin.SharePoint.Addin.Authentication;
using Microsoft.SharePoint.Client;
using System;
using System.IdentityModel.Tokens;
using System.Web;
using System.Web.Configuration;

namespace SPMVCWeb.Helpers
{
    internal static class SPContextHelper
    {
        public static RedirectionStatus CheckRedirectionStatus(HttpContextBase httpContext, out Uri redirectUrl)
        {
            return CheckRedirectionStatus(httpContext, null, out redirectUrl);
        }

        public static RedirectionStatus CheckRedirectionStatus(HttpContextBase httpContext, Uri requestUrl, out Uri redirectUrl)
        {
            if (httpContext == null)
            {
                throw new ArgumentNullException(nameof(httpContext));
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

            const string spHasRedirectedToSharePointKey = "SPHasRedirectedToSharePoint";

            if (!string.IsNullOrEmpty(httpContext.Request.QueryString[spHasRedirectedToSharePointKey]) && !contextTokenExpired)
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
            if (requestUrl != null)
            {
                var queryNameValueCollection = HttpUtility.ParseQueryString(requestUrl.Query);
                // Removes the values that are included in {StandardTokens}, as {StandardTokens} will be inserted at the beginning of the query string.
                queryNameValueCollection.Remove(SharePointContext.SPHostUrlKey);
                queryNameValueCollection.Remove(SharePointContext.SPAppWebUrlKey);
                queryNameValueCollection.Remove(SharePointContext.SPLanguageKey);
                queryNameValueCollection.Remove(SharePointContext.SPClientTagKey);
                queryNameValueCollection.Remove(SharePointContext.SPProductNumberKey);

                // Adds SPHasRedirectedToSharePoint=1.
                queryNameValueCollection.Add(spHasRedirectedToSharePointKey, "1");

                UriBuilder returnUrlBuilder = new UriBuilder(requestUrl) {Query = queryNameValueCollection.ToString()};

                // Inserts StandardTokens.
                const string standardTokens = "{StandardTokens}";
                string returnUrlString = returnUrlBuilder.Uri.AbsoluteUri;
                returnUrlString = returnUrlString.Insert(returnUrlString.IndexOf("?", StringComparison.Ordinal) + 1, standardTokens + "&");

                // Constructs redirect url.
                string redirectUrlString = TokenHelper.GetAppContextTokenRequestUrl(spHostUrl.GetLeftPart(UriPartial.Path), Uri.EscapeDataString(returnUrlString));
                redirectUrl = new Uri(redirectUrlString, UriKind.Absolute);
            }
            return RedirectionStatus.ShouldRedirect;
        }

        public static ISPContext GetSPContext(HttpContextBase httpContext)
        {
            var cookieAuthenticationEnabled = !string.IsNullOrEmpty(WebConfigurationManager.AppSettings.Get("CookieAuthenticationEnabled")) && Convert.ToBoolean(WebConfigurationManager.AppSettings.Get("CookieAuthenticationEnabled"));
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
            if (action == null) throw new ArgumentNullException(nameof(action));
            TContext clientContext = (TContext) spContext?.CreateUserClientContextForSPHost();
            if (clientContext != null)
            {
                using (clientContext)
                {
                    Action result = action.Invoke(clientContext);
                    clientContext.ExecuteQuery();
                    result?.Invoke();
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
            if (action == null) throw new ArgumentNullException(nameof(action));
            TContext clientContext = (TContext) spContext?.CreateUserClientContextForSPAppWeb();
            if (clientContext != null)
            {
                using (clientContext)
                {
                    Action result = action.Invoke(clientContext);
                    clientContext.ExecuteQuery();
                    result?.Invoke();
                }
            }
        }


        public static void ExecuteAppOnlyClientContextForSPAppWebQuery<TContext>(ISPContext spContext, Func<TContext, Action> action)
            where TContext : ClientContext
        {
            if (action == null) throw new ArgumentNullException(nameof(action));
            TContext clientContext = (TContext) spContext?.CreateAppOnlyClientContextForSPAppWeb();
            if (clientContext != null)
            {
                using (clientContext)
                {
                    Action result = action.Invoke(clientContext);
                    clientContext.ExecuteQuery();
                    result?.Invoke();
                }
            }
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
    }
}