using Microsoft.IdentityModel.S2S.Protocols.OAuth2;
using Microsoft.IdentityModel.Tokens;
using Microsoft.SharePoint.Client;
using System;
using System.Net;
using System.Security.Principal;
using System.Web;

namespace AspNet.Owin.SharePoint.Addin.Authentication
{
  /// <summary>
  /// Encapsulates all the information from SharePoint.
  /// </summary>
  public abstract class SharePointContext : ISPContext
  {
    public const string SPHostUrlKey = "SPHostUrl";
    public const string SPAppWebUrlKey = "SPAppWebUrl";
    public const string SPLanguageKey = "SPLanguage";
    public const string SPClientTagKey = "SPClientTag";
    public const string SPProductNumberKey = "SPProductNumber";

    protected static readonly TimeSpan AccessTokenLifetimeTolerance = TimeSpan.FromMinutes(5.0);

    // <AccessTokenString, UtcExpiresOn>
    protected Tuple<string, DateTime> UserAccessTokenForSpHost;
    protected Tuple<string, DateTime> UserAccessTokenForSpAppWeb;
    protected Tuple<string, DateTime> AppOnlyAccessTokenForSpHost;
    protected Tuple<string, DateTime> AppOnlyAccessTokenForSpAppWeb;

    /// <summary>
    /// Gets the SharePoint host url from QueryString of the specified HTTP request.
    /// </summary>
    /// <param name="httpRequest">The specified HTTP request.</param>
    /// <returns>The SharePoint host url. Returns <c>null</c> if the HTTP request doesn't contain the SharePoint host url.</returns>
    public static Uri GetSPHostUrl(HttpRequestBase httpRequest)
    {
      if (httpRequest == null)
      {
        throw new ArgumentNullException(nameof(httpRequest));
      }

      string spHostUrlString = TokenHelper.EnsureTrailingSlash(httpRequest.QueryString[SPHostUrlKey]);
      Uri spHostUrl;
      if (Uri.TryCreate(spHostUrlString, UriKind.Absolute, out spHostUrl) &&
          (spHostUrl.Scheme == Uri.UriSchemeHttp || spHostUrl.Scheme == Uri.UriSchemeHttps))
      {
        return spHostUrl;
      }

      return null;
    }

    /// <summary>
    /// Gets the SharePoint host url from QueryString of the specified HTTP request.
    /// </summary>
    /// <param name="httpRequest">The specified HTTP request.</param>
    /// <returns>The SharePoint host url. Returns <c>null</c> if the HTTP request doesn't contain the SharePoint host url.</returns>
    public static Uri GetSPHostUrl(HttpRequest httpRequest)
    {
      return GetSPHostUrl(new HttpRequestWrapper(httpRequest));
    }

    /// <summary>
    /// The SharePoint host url.
    /// </summary>
    public Uri SPHostUrl { get; }

    /// <summary>
    /// The SharePoint app web url.
    /// </summary>
    public Uri SPAppWebUrl { get; }

    /// <summary>
    /// The SharePoint language.
    /// </summary>
    public string SPLanguage { get; }

    /// <summary>
    /// The SharePoint client tag.
    /// </summary>
    public string SPClientTag { get; }

    /// <summary>
    /// The SharePoint product number.
    /// </summary>
    public string SPProductNumber { get; }

    /// <summary>
    /// The user access token for the SharePoint host.
    /// </summary>
    public abstract string UserAccessTokenForSPHost
    {
      get;
    }

    /// <summary>
    /// The user access token for the SharePoint app web.
    /// </summary>
    public abstract string UserAccessTokenForSPAppWeb
    {
      get;
    }

    /// <summary>
    /// The app only access token for the SharePoint host.
    /// </summary>
    public abstract string AppOnlyAccessTokenForSPHost
    {
      get;
    }

    /// <summary>
    /// The app only access token for the SharePoint app web.
    /// </summary>
    public abstract string AppOnlyAccessTokenForSPAppWeb
    {
      get;
    }

    /// <summary>
    /// Constructor.
    /// </summary>
    /// <param name="spHostUrl">The SharePoint host url.</param>
    /// <param name="spAppWebUrl">The SharePoint app web url.</param>
    /// <param name="spLanguage">The SharePoint language.</param>
    /// <param name="spClientTag">The SharePoint client tag.</param>
    /// <param name="spProductNumber">The SharePoint product number.</param>
    protected SharePointContext(Uri spHostUrl, Uri spAppWebUrl, string spLanguage, string spClientTag, string spProductNumber)
    {
      if (spHostUrl == null)
      {
        throw new ArgumentNullException(nameof(spHostUrl));
      }

      if (string.IsNullOrEmpty(spLanguage))
      {
        throw new ArgumentNullException(nameof(spLanguage));
      }

      if (string.IsNullOrEmpty(spClientTag))
      {
        throw new ArgumentNullException(nameof(spClientTag));
      }

      if (string.IsNullOrEmpty(spProductNumber))
      {
        throw new ArgumentNullException(nameof(spProductNumber));
      }

      this.SPHostUrl = spHostUrl;
      this.SPAppWebUrl = spAppWebUrl;
      this.SPLanguage = spLanguage;
      this.SPClientTag = spClientTag;
      this.SPProductNumber = spProductNumber;
    }

    /// <summary>
    /// Creates a user ClientContext for the SharePoint host.
    /// </summary>
    /// <returns>A ClientContext instance.</returns>
    public ClientContext CreateUserClientContextForSPHost()
    {
      return CreateClientContext(this.SPHostUrl, this.UserAccessTokenForSPHost);
    }

    /// <summary>
    /// Creates a user ClientContext for the SharePoint app web.
    /// </summary>
    /// <returns>A ClientContext instance.</returns>
    public ClientContext CreateUserClientContextForSPAppWeb()
    {
      return CreateClientContext(this.SPAppWebUrl, this.UserAccessTokenForSPAppWeb);
    }

    /// <summary>
    /// Creates app only ClientContext for the SharePoint host.
    /// </summary>
    /// <returns>A ClientContext instance.</returns>
    public ClientContext CreateAppOnlyClientContextForSPHost()
    {
      return CreateClientContext(this.SPHostUrl, this.AppOnlyAccessTokenForSPHost);
    }

    /// <summary>
    /// Creates an app only ClientContext for the SharePoint app web.
    /// </summary>
    /// <returns>A ClientContext instance.</returns>
    public ClientContext CreateAppOnlyClientContextForSPAppWeb()
    {
      return CreateClientContext(this.SPAppWebUrl, this.AppOnlyAccessTokenForSPAppWeb);
    }

    /// <summary>
    /// Determines if the specified access token is valid.
    /// It considers an access token as not valid if it is null, or it has expired.
    /// </summary>
    /// <param name="accessToken">The access token to verify.</param>
    /// <returns>True if the access token is valid.</returns>
    protected static bool IsAccessTokenValid(Tuple<string, DateTime> accessToken)
    {
      return accessToken != null &&
             !string.IsNullOrEmpty(accessToken.Item1) &&
             accessToken.Item2 > DateTime.UtcNow;
    }

    /// <summary>
    /// Creates a ClientContext with the specified SharePoint site url and the access token.
    /// </summary>
    /// <param name="spSiteUrl">The site url.</param>
    /// <param name="accessToken">The access token.</param>
    /// <returns>A ClientContext instance.</returns>
    private static ClientContext CreateClientContext(Uri spSiteUrl, string accessToken)
    {
      if (spSiteUrl != null && !string.IsNullOrEmpty(accessToken))
      {
        return TokenHelper.GetClientContextWithAccessToken(spSiteUrl.GetLeftPart(UriPartial.Path), accessToken);
      }
      return null;
    }
  }

  /// <summary>
  /// Redirection status.
  /// </summary>
  public enum RedirectionStatus
  {
    Ok,
    ShouldRedirect,
    CanNotRedirect
  }

  /// <summary>
  /// Provides SharePointContext instances.
  /// </summary>
  public abstract class SharePointContextProvider
  {
    private HttpContextBase _context;
    private static readonly object _lock = new object();

    /// <summary>
    /// The current SharePointContextProvider instance.
    /// </summary>
    public static SharePointContextProvider Current
    {
      get { return GetCurrentContextProvider(HttpContext.Current); }
    }

    public Uri SPHostUrl { get; set; }

    public Uri SPWebAppUrl { get; set; }

    public static SharePointContextProvider GetCurrentContextProvider(HttpContext context)
    {
      if (context == null)
      {
        throw new ArgumentNullException("context");
      }
      SharePointContextProvider wSharePointContextProvider = GetCurrentContextProvider(new HttpContextWrapper(context));
      return wSharePointContextProvider;
    }

    public static SharePointContextProvider GetCurrentContextProvider(HttpContextBase context)
    {
      if (context == null) throw new ArgumentNullException("context");
      Uri spHostUrl = null;
      if (context.Request.QueryString["SPHostUrl"] != null)
      {
        spHostUrl = SharePointContext.GetSPHostUrl(context.Request);
      }
      else
      {
        if (context.Request.Form != null && context.Request.Form.Get("SPHostUrl") != null)
        {
          var wFormCtxSpHostUrl = context.Request.Form.Get("SPHostUrl");
          spHostUrl = new Uri(HttpUtility.UrlDecode(wFormCtxSpHostUrl));
        }
      }

      string key = string.Format("DefaultSPContextProvider_{0}", spHostUrl).ToLower();
      lock (_lock)
      {
        var current = context.Items.Contains(key) ? (SharePointContextProvider)context.Items[key] : null;
        if (current == null)
        {
          if (!TokenHelper.IsHighTrustApp())
          {
            current = new SharePointAcsContextProvider();
          }
          else
          {
            current = new SharePointHighTrustContextProvider();
          }
          if (current != null)
          {
            current.SPHostUrl = spHostUrl;
            current._context = context;
            context.Items[key] = current;
          }
        }
        return current;
      }
    }

    public static String Error = "";
    /// <summary>
    /// Checks if it is necessary to redirect to SharePoint for user to authenticate.
    /// </summary>
    /// <param name="httpContext">The HTTP context.</param>
    /// <param name="redirectUrl">The redirect url to SharePoint if the status is ShouldRedirect. <c>Null</c> if the status is Ok or CanNotRedirect.</param>
    /// <returns>Redirection status.</returns>
    public static RedirectionStatus CheckRedirectionStatus(HttpContextBase httpContext, out Uri redirectUrl)
    {
      if (httpContext == null)
      {
        throw new ArgumentNullException("httpContext");
      }

      redirectUrl = null;
      Uri spHostUrl = SharePointContext.GetSPHostUrl(httpContext.Request);

      try
      {
        var spContext = GetCurrentContextProvider(httpContext).GetSharePointContext();
        if (spContext != null)
        {
          return RedirectionStatus.Ok;
        }
      }
      catch (/*SecurityTokenExpired*/Exception ex)
      {
        //Error = ex.Message;
      }

      const string SPHasRedirectedToSharePointKey = "SPHasRedirectedToSharePoint";

      if (!string.IsNullOrEmpty(httpContext.Request.QueryString[SPHasRedirectedToSharePointKey]))
      {
        // Remove 178-9808
        //Error = " Request Parameter 'SPHasRedirectedToSharePoint' is not empty (Sharepoint sent " + httpContext.Request.QueryString[SPHasRedirectedToSharePointKey] + " on " + httpContext.Request.Url.ToString();
        //return RedirectionStatus.CanNotRedirect;
      }

      if (spHostUrl == null)
      {
        Error = "SPHostUrl missing";
        return RedirectionStatus.CanNotRedirect;
      }

      Uri requestUrl = httpContext.Request.Url;

      var queryNameValueCollection = HttpUtility.ParseQueryString(requestUrl.Query);

      // Removes the values that are included in {StandardTokens}, as {StandardTokens} will be inserted at the beginning of the query string.
      queryNameValueCollection.Remove(SharePointContext.SPHostUrlKey);
      queryNameValueCollection.Remove(SharePointContext.SPAppWebUrlKey);
      queryNameValueCollection.Remove(SharePointContext.SPLanguageKey);
      queryNameValueCollection.Remove(SharePointContext.SPClientTagKey);
      queryNameValueCollection.Remove(SharePointContext.SPProductNumberKey);
      queryNameValueCollection.Remove("SPAppToken");

      // Adds SPHasRedirectedToSharePoint=1.
      queryNameValueCollection.Add(SPHasRedirectedToSharePointKey, "1");

      UriBuilder returnUrlBuilder = new UriBuilder(requestUrl);
      returnUrlBuilder.Query = queryNameValueCollection.ToString();

      // Inserts StandardTokens.
      const string StandardTokens = "{StandardTokens}";
      string returnUrlString = returnUrlBuilder.Uri.AbsoluteUri;
      returnUrlString = returnUrlString.Insert(returnUrlString.IndexOf("?") + 1, StandardTokens + "&");

      // Constructs redirect url.
      string redirectUrlString = TokenHelper.GetAppContextTokenRequestUrl(spHostUrl.AbsoluteUri, Uri.EscapeDataString(returnUrlString));

      redirectUrl = new Uri(redirectUrlString, UriKind.Absolute);

      return RedirectionStatus.ShouldRedirect;
    }

    /// <summary>
    /// Checks if it is necessary to redirect to SharePoint for user to authenticate.
    /// </summary>
    /// <param name="httpContext">The HTTP context.</param>
    /// <param name="redirectUrl">The redirect url to SharePoint if the status is ShouldRedirect. <c>Null</c> if the status is Ok or CanNotRedirect.</param>
    /// <returns>Redirection status.</returns>
    public static RedirectionStatus CheckRedirectionStatus(HttpContext httpContext, out Uri redirectUrl)
    {
      return CheckRedirectionStatus(new HttpContextWrapper(httpContext), out redirectUrl);
    }

    /// <summary>
    /// Creates a SharePointContext instance with the specified HTTP request.
    /// </summary>
    /// <param name="httpRequest">The HTTP request.</param>
    /// <returns>The SharePointContext instance. Returns <c>null</c> if errors occur.</returns>
    public SharePointContext CreateSharePointContext()
    {
      HttpRequestBase httpRequest = null;
      if (this._context != null)
      {
        httpRequest = this._context.Request;
      }
      if (httpRequest == null)
      {
        throw new ArgumentNullException("httpRequest");
      }

      // SPHostUrl
      Uri spHostUrl = SharePointContext.GetSPHostUrl(httpRequest);
      if (spHostUrl == null)
      {
        return null;
      }

      // SPAppWebUrl
      string spAppWebUrlString = TokenHelper.EnsureTrailingSlash(httpRequest.QueryString[SharePointContext.SPAppWebUrlKey]);
      if (String.IsNullOrEmpty(spAppWebUrlString)) spAppWebUrlString = HttpUtility.UrlDecode(httpRequest.QueryString["AppHostUrl"]);
      Uri spAppWebUrl;
      if (!Uri.TryCreate(spAppWebUrlString, UriKind.Absolute, out spAppWebUrl) ||
          !(spAppWebUrl.Scheme == Uri.UriSchemeHttp || spAppWebUrl.Scheme == Uri.UriSchemeHttps))
      {
        spAppWebUrl = null;
      }

      // SPLanguage
      string spLanguage = httpRequest.QueryString[SharePointContext.SPLanguageKey];
      if (string.IsNullOrEmpty(spLanguage))
      {
        return null;
      }

      // SPClientTag
      string spClientTag = httpRequest.QueryString[SharePointContext.SPClientTagKey];
      if (string.IsNullOrEmpty(spClientTag))
      {
        return null;
      }

      // SPProductNumber
      string spProductNumber = httpRequest.QueryString[SharePointContext.SPProductNumberKey];
      if (string.IsNullOrEmpty(spProductNumber))
      {
        return null;
      }

      return CreateSharePointContext(spHostUrl, spAppWebUrl, spLanguage, spClientTag, spProductNumber, httpRequest);
    }

    /// <summary>
    /// Gets a SharePointContext instance associated with the specified HTTP context.
    /// </summary>
    /// <param name="httpContext">The HTTP context.</param>
    /// <returns>The SharePointContext instance. Returns <c>null</c> if not found and a new instance can't be created.</returns>
    public SharePointContext GetSharePointContext()
    {
      var httpContext = this._context;
      if (httpContext == null)
      {
        throw new ArgumentNullException("httpContext");
      }

      Uri spHostUrl = SharePointContext.GetSPHostUrl(httpContext.Request);
      if (spHostUrl == null)
      {
        return null;
      }

      SharePointContext spContext = LoadSharePointContext(httpContext);
      if (spContext == null || !ValidateSharePointContext(spContext, httpContext, spHostUrl))
      {
        spContext = CreateSharePointContext();
        if (spContext != null)
        {
          SaveSharePointContext(spContext, httpContext);
        }
      }
      return spContext;
    }

    /// <summary>
    /// Creates a SharePointContext instance.
    /// </summary>
    /// <param name="spHostUrl">The SharePoint host url.</param>
    /// <param name="spAppWebUrl">The SharePoint app web url.</param>
    /// <param name="spLanguage">The SharePoint language.</param>
    /// <param name="spClientTag">The SharePoint client tag.</param>
    /// <param name="spProductNumber">The SharePoint product number.</param>
    /// <param name="httpRequest">The HTTP request.</param>
    /// <returns>The SharePointContext instance. Returns <c>null</c> if errors occur.</returns>
    protected abstract SharePointContext CreateSharePointContext(Uri spHostUrl, Uri spAppWebUrl, string spLanguage, string spClientTag, string spProductNumber, HttpRequestBase httpRequest);

    /// <summary>
    /// Validates if the given SharePointContext can be used with the specified HTTP context.
    /// </summary>
    /// <param name="spContext">The SharePointContext.</param>
    /// <param name="httpContext">The HTTP context.</param>
    /// <returns>True if the given SharePointContext can be used with the specified HTTP context.</returns>
    protected abstract bool ValidateSharePointContext(SharePointContext spContext, HttpContextBase httpContext, Uri requestUrl);

    /// <summary>
    /// Loads the SharePointContext instance associated with the specified HTTP context.
    /// </summary>
    /// <param name="httpContext">The HTTP context.</param>
    /// <returns>The SharePointContext instance. Returns <c>null</c> if not found.</returns>
    protected abstract SharePointContext LoadSharePointContext(HttpContextBase httpContext);

    /// <summary>
    /// Saves the specified SharePointContext instance associated with the specified HTTP context.
    /// <c>null</c> is accepted for clearing the SharePointContext instance associated with the HTTP context.
    /// </summary>
    /// <param name="spContext">The SharePointContext instance to be saved, or <c>null</c>.</param>
    /// <param name="httpContext">The HTTP context.</param>
    protected abstract void SaveSharePointContext(SharePointContext spContext, HttpContextBase httpContext);

    protected abstract void Remove(HttpContextBase httpContext);
  }

  #region ACS

  /// <summary>
  /// Encapsulates all the information from SharePoint in ACS mode.
  /// </summary>
  public class SharePointAcsContext : SharePointContext
  {
    private readonly string _contextToken;
    private readonly SharePointContextToken _contextTokenObj;

    /// <summary>
    /// The context token.
    /// </summary>
    public string ContextToken => this._contextTokenObj.ValidTo > DateTime.UtcNow ? this._contextToken : null;

    /// <summary>
    /// The context token's "CacheKey" claim.
    /// </summary>
    public string CacheKey => this._contextTokenObj.ValidTo > DateTime.UtcNow ? this._contextTokenObj.CacheKey : null;

    /// <summary>
    /// The context token's "refreshtoken" claim.
    /// </summary>
    public string RefreshToken => this._contextTokenObj.ValidTo > DateTime.UtcNow ? this._contextTokenObj.RefreshToken : null;

    public override string UserAccessTokenForSPHost
    {
      get
      {
        return GetAccessTokenString(ref this.UserAccessTokenForSpHost,
                                    () =>
        //TokenHelper.GetAccessToken(this._contextTokenObj.RefreshToken, TokenHelper.SharePointPrincipal, this.SPHostUrl.Authority, TokenHelper.GetRealmFromTargetUrl(this.SPHostUrl))
        TokenHelper.GetAccessToken(this._contextTokenObj, this.SPHostUrl.Authority));
      }
    }

    public override string UserAccessTokenForSPAppWeb
    {
      get
      {
        if (this.SPAppWebUrl == null)
        {
          return null;
        }

        return GetAccessTokenString(ref this.UserAccessTokenForSpAppWeb,
                                    () => TokenHelper.GetAccessToken(this._contextTokenObj, this.SPAppWebUrl.Authority));
      }
    }

    public override string AppOnlyAccessTokenForSPHost
    {
      get
      {
        return GetAccessTokenString(ref this.AppOnlyAccessTokenForSpHost,
                                    () => TokenHelper.GetAppOnlyAccessToken(TokenHelper.SharePointPrincipal, this.SPHostUrl.Authority, TokenHelper.GetRealmFromTargetUrl(this.SPHostUrl)));
      }
    }

    public override string AppOnlyAccessTokenForSPAppWeb
    {
      get
      {
        if (this.SPAppWebUrl == null)
        {
          return null;
        }

        return GetAccessTokenString(ref this.AppOnlyAccessTokenForSpAppWeb,
                                    () => TokenHelper.GetAppOnlyAccessToken(TokenHelper.SharePointPrincipal, this.SPAppWebUrl.Authority, TokenHelper.GetRealmFromTargetUrl(this.SPAppWebUrl)));
      }
    }

    public SharePointAcsContext(Uri spHostUrl, Uri spAppWebUrl, string spLanguage, string spClientTag, string spProductNumber, string contextToken, SharePointContextToken contextTokenObj)
        : base(spHostUrl, spAppWebUrl, spLanguage, spClientTag, spProductNumber)
    {
      if (string.IsNullOrEmpty(contextToken))
      {
        throw new ArgumentNullException(nameof(contextToken));
      }

      if (contextTokenObj == null)
      {
        throw new ArgumentNullException(nameof(contextTokenObj));
      }

      this._contextToken = contextToken;
      this._contextTokenObj = contextTokenObj;
    }

    /// <summary>
    /// Ensures the access token is valid and returns it.
    /// </summary>
    /// <param name="accessToken">The access token to verify.</param>
    /// <param name="tokenRenewalHandler">The token renewal handler.</param>
    /// <returns>The access token string.</returns>
    private static string GetAccessTokenString(ref Tuple<string, DateTime> accessToken, Func<OAuth2AccessTokenResponse> tokenRenewalHandler)
    {
      RenewAccessTokenIfNeeded(ref accessToken, tokenRenewalHandler);

      return IsAccessTokenValid(accessToken) ? accessToken.Item1 : null;
    }

    /// <summary>
    /// Renews the access token if it is not valid.
    /// </summary>
    /// <param name="accessToken">The access token to renew.</param>
    /// <param name="tokenRenewalHandler">The token renewal handler.</param>
    private static void RenewAccessTokenIfNeeded(ref Tuple<string, DateTime> accessToken, Func<OAuth2AccessTokenResponse> tokenRenewalHandler)
    {
      if (IsAccessTokenValid(accessToken))
      {
        return;
      }

      try
      {
        OAuth2AccessTokenResponse oAuth2AccessTokenResponse = tokenRenewalHandler();

        DateTime expiresOn = oAuth2AccessTokenResponse.ExpiresOn;

        if ((expiresOn - oAuth2AccessTokenResponse.NotBefore) > AccessTokenLifetimeTolerance)
        {
          // Make the access token get renewed a bit earlier than the time when it expires
          // so that the calls to SharePoint with it will have enough time to complete successfully.
          expiresOn -= AccessTokenLifetimeTolerance;
        }

        accessToken = Tuple.Create(oAuth2AccessTokenResponse.AccessToken, expiresOn);
      }
      catch (WebException)
      {
      }
    }
  }

  /// <summary>
  /// Default provider for SharePointAcsContext.
  /// </summary>
  public class SharePointAcsContextProvider : SharePointContextProvider
  {
    private const string SPContextKey = "SPContext";
    private const string SPCacheKeyKey = "SPCacheKey";

    protected override SharePointContext CreateSharePointContext(Uri spHostUrl, Uri spAppWebUrl, string spLanguage, string spClientTag, string spProductNumber, HttpRequestBase httpRequest)
    {
      string contextTokenString = TokenHelper.GetContextTokenFromRequest(httpRequest);
      if (string.IsNullOrEmpty(contextTokenString))
      {
        return null;
      }

      SharePointContextToken contextToken = null;
      try
      {
        contextToken = TokenHelper.ReadAndValidateContextToken(contextTokenString, httpRequest.Url.Authority);
      }
      catch (WebException)
      {
        return null;
      }
      catch (AudienceUriValidationFailedException)
      {
        return null;
      }

      return new SharePointAcsContext(spHostUrl, spAppWebUrl, spLanguage, spClientTag, spProductNumber, contextTokenString, contextToken);
    }

    protected override bool ValidateSharePointContext(SharePointContext spContext, HttpContextBase httpContext, Uri spHostUrl)
    {
      SharePointAcsContext spAcsContext = spContext as SharePointAcsContext;

      if (spAcsContext != null)
      {
        //Uri spHostUrl = SharePointContext.GetSPHostUrl(httpContext.Request);
        string contextToken = TokenHelper.GetContextTokenFromRequest(httpContext.Request);
        HttpCookie spCacheKeyCookie = httpContext.Request.Cookies[SPCacheKeyKey];
        string spCacheKey = spCacheKeyCookie != null ? spCacheKeyCookie.Value : null;

        return string.Equals(spHostUrl.GetLeftPart(UriPartial.Path).TrimEnd('/'), spAcsContext.SPHostUrl.GetLeftPart(UriPartial.Path).TrimEnd('/'), StringComparison.OrdinalIgnoreCase) &&
               !string.IsNullOrEmpty(spAcsContext.CacheKey) &&
               spCacheKey == spAcsContext.CacheKey &&
               !string.IsNullOrEmpty(spAcsContext.ContextToken) &&
               (string.IsNullOrEmpty(contextToken) || contextToken == spAcsContext.ContextToken);
      }

      return false;
    }

    protected override SharePointContext LoadSharePointContext(HttpContextBase httpContext)
    {
      return httpContext.Session?[SPContextKey] as SharePointAcsContext;
    }

    protected override void SaveSharePointContext(SharePointContext spContext, HttpContextBase httpContext)
    {
      SharePointAcsContext spAcsContext = spContext as SharePointAcsContext;

      if (spAcsContext != null)
      {
        HttpCookie spCacheKeyCookie = new HttpCookie(SPCacheKeyKey)
        {
          Value = spAcsContext.CacheKey,
          Secure = true,
          HttpOnly = true
        };

        httpContext.Response.AppendCookie(spCacheKeyCookie);
      }

      if (httpContext.Session != null)
      {
        httpContext.Session[SPContextKey] = spAcsContext;
      }
    }

    protected override void Remove(HttpContextBase httpContext)
    {
      HttpCookie spCacheKeyCookie = httpContext.Request.Cookies[SPCacheKeyKey];
      if (spCacheKeyCookie != null)
      {
        spCacheKeyCookie.Expires = DateTime.Now.AddDays(-1);
      }
      if (httpContext.Session != null)
      {
        httpContext.Session.Remove(SPContextKey);
        httpContext.Session.Abandon();
      }
    }
  }

  #endregion ACS

  #region HighTrust

  /// <summary>
  /// Encapsulates all the information from SharePoint in HighTrust mode.
  /// </summary>
  public class SharePointHighTrustContext : SharePointContext
  {
    /// <summary>
    /// The Windows identity for the current user.
    /// </summary>
    public WindowsIdentity LogonUserIdentity { get; }

    public override string UserAccessTokenForSPHost
    {
      get
      {
        return GetAccessTokenString(ref this.UserAccessTokenForSpHost,
                                    () => TokenHelper.GetS2SAccessTokenWithWindowsIdentity(this.SPHostUrl, this.LogonUserIdentity));
      }
    }

    public override string UserAccessTokenForSPAppWeb
    {
      get
      {
        if (this.SPAppWebUrl == null)
        {
          return null;
        }

        return GetAccessTokenString(ref this.UserAccessTokenForSpAppWeb,
                                    () => TokenHelper.GetS2SAccessTokenWithWindowsIdentity(this.SPAppWebUrl, this.LogonUserIdentity));
      }
    }

    public override string AppOnlyAccessTokenForSPHost
    {
      get
      {
        return GetAccessTokenString(ref this.AppOnlyAccessTokenForSpHost,
                                    () => TokenHelper.GetS2SAccessTokenWithWindowsIdentity(this.SPHostUrl, null));
      }
    }

    public override string AppOnlyAccessTokenForSPAppWeb
    {
      get
      {
        if (this.SPAppWebUrl == null)
        {
          return null;
        }

        return GetAccessTokenString(ref this.AppOnlyAccessTokenForSpAppWeb,
                                    () => TokenHelper.GetS2SAccessTokenWithWindowsIdentity(this.SPAppWebUrl, null));
      }
    }

    public SharePointHighTrustContext(Uri spHostUrl, Uri spAppWebUrl, string spLanguage, string spClientTag, string spProductNumber, WindowsIdentity logonUserIdentity)
        : base(spHostUrl, spAppWebUrl, spLanguage, spClientTag, spProductNumber)
    {
      if (logonUserIdentity == null)
      {
        throw new ArgumentNullException(nameof(logonUserIdentity));
      }

      this.LogonUserIdentity = logonUserIdentity;
    }

    /// <summary>
    /// Ensures the access token is valid and returns it.
    /// </summary>
    /// <param name="accessToken">The access token to verify.</param>
    /// <param name="tokenRenewalHandler">The token renewal handler.</param>
    /// <returns>The access token string.</returns>
    private static string GetAccessTokenString(ref Tuple<string, DateTime> accessToken, Func<string> tokenRenewalHandler)
    {
      RenewAccessTokenIfNeeded(ref accessToken, tokenRenewalHandler);

      return IsAccessTokenValid(accessToken) ? accessToken.Item1 : null;
    }

    /// <summary>
    /// Renews the access token if it is not valid.
    /// </summary>
    /// <param name="accessToken">The access token to renew.</param>
    /// <param name="tokenRenewalHandler">The token renewal handler.</param>
    private static void RenewAccessTokenIfNeeded(ref Tuple<string, DateTime> accessToken, Func<string> tokenRenewalHandler)
    {
      if (IsAccessTokenValid(accessToken))
      {
        return;
      }

      DateTime expiresOn = DateTime.UtcNow.Add(TokenHelper.HighTrustAccessTokenLifetime);

      if (TokenHelper.HighTrustAccessTokenLifetime > AccessTokenLifetimeTolerance)
      {
        // Make the access token get renewed a bit earlier than the time when it expires
        // so that the calls to SharePoint with it will have enough time to complete successfully.
        expiresOn -= AccessTokenLifetimeTolerance;
      }

      accessToken = Tuple.Create(tokenRenewalHandler(), expiresOn);
    }
  }

  /// <summary>
  /// Default provider for SharePointHighTrustContext.
  /// </summary>
  public class SharePointHighTrustContextProvider : SharePointContextProvider
  {
    private const string SPContextKey = "SPContext";

    protected override SharePointContext CreateSharePointContext(Uri spHostUrl, Uri spAppWebUrl, string spLanguage, string spClientTag, string spProductNumber, HttpRequestBase httpRequest)
    {
      WindowsIdentity logonUserIdentity = httpRequest.LogonUserIdentity;
      if (logonUserIdentity == null || !logonUserIdentity.IsAuthenticated || logonUserIdentity.IsGuest || logonUserIdentity.User == null)
      {
        return null;
      }
      return new SharePointHighTrustContext(spHostUrl, spAppWebUrl, spLanguage, spClientTag, spProductNumber, logonUserIdentity);
    }

    protected override bool ValidateSharePointContext(SharePointContext spContext, HttpContextBase httpContext, Uri spHostUrl)
    {
      SharePointHighTrustContext spHighTrustContext = spContext as SharePointHighTrustContext;

      if (spHighTrustContext != null)
      {
        if (spHostUrl == null)
        {
          spHostUrl = SharePointContext.GetSPHostUrl(httpContext.Request);
        }
        WindowsIdentity logonUserIdentity = httpContext.Request.LogonUserIdentity;
        return string.Equals(spHostUrl.GetLeftPart(UriPartial.Path).TrimEnd('/'), spHighTrustContext.SPHostUrl.GetLeftPart(UriPartial.Path).TrimEnd('/'), StringComparison.OrdinalIgnoreCase) &&
               logonUserIdentity != null &&
               logonUserIdentity.IsAuthenticated &&
               !logonUserIdentity.IsGuest &&
               logonUserIdentity.User == spHighTrustContext.LogonUserIdentity.User;
      }

      return false;
    }

    protected override SharePointContext LoadSharePointContext(HttpContextBase httpContext)
    {
      return httpContext.Session?[SPContextKey] as SharePointHighTrustContext;
    }

    protected override void SaveSharePointContext(SharePointContext spContext, HttpContextBase httpContext)
    {
      if (httpContext.Session != null)
      {
        httpContext.Session[SPContextKey] = spContext as SharePointHighTrustContext;
      }
    }

    protected override void Remove(HttpContextBase httpContext)
    {
      if (httpContext.Session != null)
      {
        httpContext.Session.Remove(SPContextKey);
        httpContext.Session.Abandon();
      }
    }
  }

  #endregion HighTrust
}
