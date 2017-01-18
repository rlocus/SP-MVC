using System.Linq;
using System.Security.Claims;

namespace AspNet.Owin.SharePoint.Addin.Authentication
{
    static class ClaimsIdentityExtensions
    {
        public static void SetClaim(this ClaimsIdentity identity, Claim claim)
        {
            var existingClaims = identity.Claims.Where(c => c.Type == claim.Type).ToList();
            if (existingClaims.Count > 0)
            {
                foreach (var existingClaim in existingClaims)
                {
                    identity.RemoveClaim(existingClaim);
                }
            }
            identity.AddClaim(claim);
        }
    }
}