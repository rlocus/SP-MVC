using System.Security.Claims;

namespace AspNet.Owin.SharePoint.Addin.Authentication
{
	public static class SPContextProvider
	{
		public static SPContext Get(ClaimsPrincipal claimsPrincipal)
		{
		    return Get((ClaimsIdentity)claimsPrincipal.Identity);
		}

        public static SPContext Get(ClaimsIdentity claimsIdentity)
        {
            if (!TokenHelper.IsHighTrustApp())
            {
                return new AcsContext(claimsIdentity);
            }
            return new HighTrustContext(claimsIdentity);
        }
    }
}
