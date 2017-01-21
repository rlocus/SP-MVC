using System;
using System.Security.Claims;

namespace AspNet.Owin.SharePoint.Addin.Authentication
{
    public class AcsContext : SPContext
    {
        public AcsContext(ClaimsPrincipal claimsPrincipal) : base(claimsPrincipal)
        {
        }

        public AcsContext(ClaimsIdentity claimsIdentity) : base(claimsIdentity)
        {
        }

        protected override AccessToken CreateUserAccessToken(Uri host)
        {
            var token = TokenHelper.GetAccessToken(RefreshToken, TargetPrincipalName, host.Authority, Realm);
            //ClaimsIdentity.SetClaim(new Claim(SPAddinClaimTypes.RefreshToken, token.RefreshToken));
            DateTime expiresOn = token.ExpiresOn;
            if ((expiresOn - token.NotBefore) > AccessTokenLifetimeTolerance)
            {
                // Make the access token get renewed a bit earlier than the time when it expires
                // so that the calls to SharePoint with it will have enough time to complete successfully.
                expiresOn -= AccessTokenLifetimeTolerance;
            }
            return new AccessToken
            {
                Value = token.AccessToken,
                ExpiresOn = expiresOn
            };
        }

        protected override AccessToken CreateAppOnlyAccessToken(Uri host)
        {
            var oauthToken = TokenHelper.GetAppOnlyAccessToken(TokenHelper.SharePointPrincipal, host.Authority, Realm);
            DateTime expiresOn = oauthToken.ExpiresOn;
            if ((expiresOn - oauthToken.NotBefore) > AccessTokenLifetimeTolerance)
            {
                // Make the access token get renewed a bit earlier than the time when it expires
                // so that the calls to SharePoint with it will have enough time to complete successfully.
                expiresOn -= AccessTokenLifetimeTolerance;
            }
            return new AccessToken
            {
                Value = oauthToken.AccessToken,
                ExpiresOn = expiresOn
            };
        }
    }
}
