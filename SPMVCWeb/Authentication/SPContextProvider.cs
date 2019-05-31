using System.Security.Claims;

namespace AspNet.Owin.SharePoint.Addin.Authentication
{
  public static class SPContextProvider
  {
    public static SPContext Get(ClaimsPrincipal claimsPrincipal, bool isWebPart)
    {
      return Get((ClaimsIdentity)claimsPrincipal.Identity, isWebPart);
    }

    public static SPContext Get(ClaimsIdentity claimsIdentity, bool isWebPart)
    {
      if (!TokenHelper.IsHighTrustApp())
      {
        return new AcsContext(claimsIdentity, isWebPart);
      }
      return new HighTrustContext(claimsIdentity, isWebPart);
    }
  }
}
