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
			var oauthToken = TokenHelper.GetAccessToken(RefreshToken, TargetPrincipalName, host.Authority, Realm);
			return new AccessToken
			{
				Value = oauthToken.AccessToken,
				ExpiredOn = oauthToken.ExpiresOn.AddMinutes(-5)
			};
		}

		protected override AccessToken CreateAppOnlyAccessToken(Uri host)
		{
			var oauthToken = TokenHelper.GetAppOnlyAccessToken(TokenHelper.SharePointPrincipal, host.Authority, Realm);

			return new AccessToken
			{
				Value = oauthToken.AccessToken,
				ExpiredOn = oauthToken.ExpiresOn.AddMinutes(-5)
			};
		}
	}
}
