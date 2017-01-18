using System;
using System.Security.Claims;

namespace AspNet.Owin.SharePoint.Addin.Authentication
{
    public class HighTrustContext : SPContext
    {
        private readonly string _userId;

        public HighTrustContext(ClaimsPrincipal claimsPrincipal) : base(claimsPrincipal)
        {
            var userIdClaim = claimsPrincipal.FindFirst(c => c.Type.Equals(SPAddinClaimTypes.ADUserId));
            if (userIdClaim != null) _userId = userIdClaim.Value;
        }

        public HighTrustContext(ClaimsIdentity claimsIdentity) : base(claimsIdentity)
        {
            var userIdClaim = claimsIdentity.FindFirst(c => c.Type.Equals(SPAddinClaimTypes.ADUserId));
            if (userIdClaim != null) _userId = userIdClaim.Value;
        }

        protected override AccessToken CreateAppOnlyAccessToken(Uri host)
        {
            var s2sToken = OwinTokenHelper.GetS2SAccessToken(host, null);
            DateTime expiresOn = DateTime.UtcNow.Add(TokenHelper.HighTrustAccessTokenLifetime);
            expiresOn -= AccessTokenLifetimeTolerance;
            return new AccessToken
            {
                Value = s2sToken,
                ExpiresOn = expiresOn
            };
        }

        protected override AccessToken CreateUserAccessToken(Uri host)
        {
            var s2sToken = OwinTokenHelper.GetS2SAccessToken(host, _userId);
            DateTime expiresOn = DateTime.UtcNow.Add(TokenHelper.HighTrustAccessTokenLifetime);
            expiresOn -= AccessTokenLifetimeTolerance;
            return new AccessToken
            {
                Value = s2sToken,
                ExpiresOn = expiresOn
            };
        }
    }
}
