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

        ClientContext CreateUserClientContextForSPHost();
        ClientContext CreateUserClientContextForSPAppWeb();

        ClientContext CreateAppOnlyClientContextForSPHost();
        ClientContext CreateAppOnlyClientContextForSPAppWeb();
    }
    public abstract class SPContext : ISPContext
    {
        protected static readonly ITokenCache Cache;

        protected readonly ClaimsIdentity _claimsIdentity;

        protected static readonly TimeSpan AccessTokenLifetimeTolerance = TimeSpan.FromMinutes(5.0);

        public Guid ClientId
        {
            get
            {
                if (!_claimsIdentity.HasClaim(c => c.Type == SPAddinClaimTypes.ClientId))
                {
                    return Guid.Empty;
                    //throw new Exception("Unable to find Client Id under current user's claims");
                }
                string clientId = _claimsIdentity.FindFirst(SPAddinClaimTypes.ClientId).Value;
                return string.IsNullOrEmpty(clientId) ? Guid.Empty : new Guid(clientId);
            }
        }

        public string RefreshToken
        {
            get
            {
                if (!_claimsIdentity.HasClaim(c => c.Type == SPAddinClaimTypes.RefreshToken))
                {
                    return null;
                    //throw new Exception("Unable to find Refresh Token under current user's claims");
                }

                return _claimsIdentity.FindFirst(SPAddinClaimTypes.RefreshToken).Value;
            }
        }

        public string UserKey
        {
            get
            {
                if (!_claimsIdentity.HasClaim(c => c.Type == SPAddinClaimTypes.CacheKey))
                {
                    //throw new Exception("Unable to find User Hash Key under current user's claims");
                    return null;
                }

                return _claimsIdentity.FindFirst(SPAddinClaimTypes.CacheKey).Value;
            }
        }

        public string Realm
        {
            get
            {
                if (!_claimsIdentity.HasClaim(c => c.Type == SPAddinClaimTypes.Realm))
                {
                    return null;
                    //throw new Exception("Unable to find Realm under current user's claims");
                }

                return _claimsIdentity.FindFirst(SPAddinClaimTypes.Realm).Value;
            }
        }

        public string TargetPrincipalName
        {
            get
            {
                if (!_claimsIdentity.HasClaim(c => c.Type == SPAddinClaimTypes.TargetPrincipalName))
                {
                    return null;
                    //throw new Exception("Unable to find TargetPrincipalName under current user's claims");
                }

                return _claimsIdentity.FindFirst(SPAddinClaimTypes.TargetPrincipalName).Value;
            }
        }

        public Uri SPHostUrl
        {
            get
            {
                if (!_claimsIdentity.HasClaim(c => c.Type == SPAddinClaimTypes.SPHostUrl))
                {
                    return null;
                    //throw new Exception("Unable to find SPHostUrl under current user's claims");
                }

                return new Uri(_claimsIdentity.FindFirst(SPAddinClaimTypes.SPHostUrl).Value);
            }
        }

        public Uri SPAppWebUrl
        {
            get
            {
                if (!_claimsIdentity.HasClaim(c => c.Type == SPAddinClaimTypes.SPAppWebUrl))
                {
                    return null;
                    //throw new Exception("Unable to find SPAppWebUrl under current user's claims");
                }

                return new Uri(_claimsIdentity.FindFirst(SPAddinClaimTypes.SPAppWebUrl).Value);
            }
        }

        static SPContext()
        {
            Cache = new DefaultTokenCache();
        }

        protected SPContext(ClaimsPrincipal claimsPrincipal)
        {
            if (claimsPrincipal == null) throw new ArgumentNullException("claimsPrincipal");
            _claimsIdentity = (ClaimsIdentity)claimsPrincipal.Identity;
        }

        protected SPContext(ClaimsIdentity claimsIdentity)
        {
            if (claimsIdentity == null) throw new ArgumentNullException("claimsIdentity");
            _claimsIdentity = claimsIdentity;
        }

        protected ClientContext GetUserClientContext(Uri host)
        {
            var accessToken = Cache.Get(GetUserCacheKey(host.Authority));
            if (accessToken == null)
            {
                accessToken = CreateUserAccessToken(host);
                Cache.Insert(accessToken, GetUserCacheKey(host.Authority));
            }
            return TokenHelper.GetClientContextWithAccessToken(host.GetLeftPart(UriPartial.Path), accessToken.Value);
        }

        protected ClientContext GetAppOnlyClientContext(Uri host)
        {
            var accessToken = Cache.Get(GetAppOnlyCacheKey(host.Authority));
            if (accessToken == null)
            {
                accessToken = CreateAppOnlyAccessToken(host);
                Cache.Insert(accessToken, GetAppOnlyCacheKey(host.Authority));
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

        //protected bool IsAccessTokenValid(AccessToken token)
        //{
        //	return !string.IsNullOrEmpty(token?.Value) && token.ExpiredOn > DateTime.UtcNow;
        //}

        public void ClearCache()
        {
            if (this.SPHostUrl != null)
            {
                Cache.Remove(GetAppOnlyCacheKey(this.SPHostUrl.Authority));
                Cache.Remove(GetUserCacheKey(this.SPHostUrl.Authority));
            }
        }
    }
}
