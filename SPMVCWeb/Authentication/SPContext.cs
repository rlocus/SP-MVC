using System;
using System.Security.Claims;
using AspNet.Owin.SharePoint.Addin.Authentication.Caching;
using Microsoft.SharePoint.Client;

namespace AspNet.Owin.SharePoint.Addin.Authentication
{
  interface ISPContext
  {
    Uri SPHostUrl { get; }

    Uri SPAppWebUrl { get; }

    bool IsWebPart { get; }

    ClientContext CreateUserClientContextForSPHost();
    ClientContext CreateUserClientContextForSPAppWeb();

    ClientContext CreateAppOnlyClientContextForSPHost();
    ClientContext CreateAppOnlyClientContextForSPAppWeb();
  }
  public abstract class SPContext : ISPContext
  {
    protected static readonly ITokenCache Cache;

    protected readonly ClaimsIdentity ClaimsIdentity;

    protected static readonly TimeSpan AccessTokenLifetimeTolerance = TimeSpan.FromMinutes(5.0);

    public Guid ClientId
    {
      get
      {
        if (!ClaimsIdentity.HasClaim(c => c.Type == SPAddinClaimTypes.ClientId))
        {
          return Guid.Empty;
          //throw new Exception("Unable to find Client Id under current user's claims");
        }
        string clientId = ClaimsIdentity.FindFirst(SPAddinClaimTypes.ClientId).Value;
        return string.IsNullOrEmpty(clientId) ? Guid.Empty : new Guid(clientId);
      }
    }

    public string RefreshToken
    {
      get
      {
        if (!ClaimsIdentity.HasClaim(c => c.Type == SPAddinClaimTypes.RefreshToken))
        {
          return null;
          //throw new Exception("Unable to find Refresh Token under current user's claims");
        }

        return ClaimsIdentity.FindFirst(SPAddinClaimTypes.RefreshToken).Value;
      }
    }

    public string UserKey
    {
      get
      {
        if (!ClaimsIdentity.HasClaim(c => c.Type == SPAddinClaimTypes.CacheKey))
        {
          //throw new Exception("Unable to find User Hash Key under current user's claims");
          return null;
        }

        return ClaimsIdentity.FindFirst(SPAddinClaimTypes.CacheKey).Value;
      }
    }

    public string Realm
    {
      get
      {
        if (!ClaimsIdentity.HasClaim(c => c.Type == SPAddinClaimTypes.Realm))
        {
          return null;
          //throw new Exception("Unable to find Realm under current user's claims");
        }

        return ClaimsIdentity.FindFirst(SPAddinClaimTypes.Realm).Value;
      }
    }

    public string TargetPrincipalName
    {
      get
      {
        if (!ClaimsIdentity.HasClaim(c => c.Type == SPAddinClaimTypes.TargetPrincipalName))
        {
          return null;
          //throw new Exception("Unable to find TargetPrincipalName under current user's claims");
        }

        return ClaimsIdentity.FindFirst(SPAddinClaimTypes.TargetPrincipalName).Value;
      }
    }

    public Uri SPHostUrl
    {
      get
      {
        if (!ClaimsIdentity.HasClaim(c => c.Type == SPAddinClaimTypes.SPHostUrl))
        {
          return null;
          //throw new Exception("Unable to find SPHostUrl under current user's claims");
        }

        return new Uri(ClaimsIdentity.FindFirst(SPAddinClaimTypes.SPHostUrl).Value);
      }
    }

    public Uri SPAppWebUrl
    {
      get
      {
        if (!ClaimsIdentity.HasClaim(c => c.Type == SPAddinClaimTypes.SPAppWebUrl))
        {
          return null;
          //throw new Exception("Unable to find SPAppWebUrl under current user's claims");
        }

        return new Uri(ClaimsIdentity.FindFirst(SPAddinClaimTypes.SPAppWebUrl).Value);
      }
    }

    public bool IsWebPart { get; }

    static SPContext()
    {
      Cache = new DefaultTokenCache();
    }

    protected SPContext(ClaimsPrincipal claimsPrincipal, bool isWebPart)
    {
      if (claimsPrincipal == null) throw new ArgumentNullException(nameof(claimsPrincipal));
      ClaimsIdentity = (ClaimsIdentity)claimsPrincipal.Identity;
      IsWebPart = isWebPart;
    }

    protected SPContext(ClaimsIdentity claimsIdentity, bool isWebPart)
    {
      ClaimsIdentity = claimsIdentity ?? throw new ArgumentNullException(nameof(claimsIdentity));
      IsWebPart = isWebPart;
    }

    protected ClientContext GetUserClientContext(Uri host)
    {
      string cacheKey = GetUserCacheKey(host.Authority);
      AccessToken accessToken = Cache.Get(cacheKey);
      if (accessToken == null || !accessToken.IsValid())
      {
        accessToken = CreateUserAccessToken(host);
        Cache.Insert(accessToken, cacheKey);
      }
      return TokenHelper.GetClientContextWithAccessToken(host.GetLeftPart(UriPartial.Path), accessToken.Value);
    }

    protected ClientContext GetAppOnlyClientContext(Uri host)
    {
      string cacheKey = GetAppOnlyCacheKey(host.Authority);
      AccessToken accessToken = Cache.Get(cacheKey);
      if (accessToken == null || !accessToken.IsValid())
      {
        accessToken = CreateAppOnlyAccessToken(host);
        Cache.Insert(accessToken, cacheKey);
      }
      return TokenHelper.GetClientContextWithAccessToken(host.GetLeftPart(UriPartial.Path), accessToken.Value);
    }

    public ClientContext CreateUserClientContextForSPHost()
    {
      return GetUserClientContext(SPHostUrl);
    }

    public ClientContext CreateUserClientContextForSPAppWeb()
    {
      return GetUserClientContext(SPAppWebUrl);
    }

    public ClientContext CreateAppOnlyClientContextForSPHost()
    {
      return GetAppOnlyClientContext(SPHostUrl);
    }

    public ClientContext CreateAppOnlyClientContextForSPAppWeb()
    {
      return GetAppOnlyClientContext(SPAppWebUrl);
    }

    public AccessToken CreateUserAccessTokenForSPHost()
    {
      return CreateUserAccessToken(SPHostUrl);
    }

    public AccessToken CreateUserAccessTokenSPAppWeb()
    {
      return CreateUserAccessToken(SPAppWebUrl);
    }

    public AccessToken CreateAppOnlyAccessTokenForSPHost()
    {
      return CreateAppOnlyAccessToken(SPHostUrl);
    }

    public AccessToken CreateAppOnlyAccessTokenSPAppWeb()
    {
      return CreateAppOnlyAccessToken(SPAppWebUrl);
    }

    protected abstract AccessToken CreateAppOnlyAccessToken(Uri host);
    protected abstract AccessToken CreateUserAccessToken(Uri host);

    protected string GetUserCacheKey(string host)
    {
      return $"{UserKey}_{host}";
    }

    protected string GetAppOnlyCacheKey(string host)
    {
      return $"{Realm}_{host}";
    }
  }
}
