using System;
using System.Security.Claims;

namespace AspNet.Owin.SharePoint.Addin.Authentication
{
    public class HighTrustContext : SPContext
    {
        private readonly string _userId;

        public HighTrustContext(ClaimsPrincipal claimsPrincipal, bool isWebPart) : base(claimsPrincipal, isWebPart)
        {
            var userIdClaim = claimsPrincipal.FindFirst(c => c.Type.Equals(SPAddinClaimTypes.ADUserId));
            if (userIdClaim != null) _userId = userIdClaim.Value;
        }

        public HighTrustContext(ClaimsIdentity claimsIdentity, bool isWebPart) : base(claimsIdentity, isWebPart)
        {
            var userIdClaim = claimsIdentity.FindFirst(c => c.Type.Equals(SPAddinClaimTypes.ADUserId));
            if (userIdClaim != null) _userId = userIdClaim.Value;
        }

        protected override AccessToken CreateAppOnlyAccessToken(Uri host)
        {
            var token = OwinTokenHelper.GetS2SAccessToken(host, null);
            DateTime expiresOn = DateTime.UtcNow.Add(TokenHelper.HighTrustAccessTokenLifetime);
            expiresOn -= AccessTokenLifetimeTolerance;
            return new AccessToken
            {
                Value = token,
                ExpiresOn = expiresOn
            };
        }

        protected override AccessToken CreateUserAccessToken(Uri host)
        {
            var token = OwinTokenHelper.GetS2SAccessToken(host, _userId);
            DateTime expiresOn = DateTime.UtcNow.Add(TokenHelper.HighTrustAccessTokenLifetime);
            expiresOn -= AccessTokenLifetimeTolerance;
            return new AccessToken
            {
                Value = token,
                ExpiresOn = expiresOn
            };
        }
    }
}
