using System;
using System.Net;
using System.Security.Principal;
using System.Threading.Tasks;
using AspNet.Owin.SharePoint.Addin.Authentication.Provider;
using Microsoft.Owin.Logging;
using Microsoft.Owin.Security;
using Microsoft.Owin.Security.Infrastructure;
using Microsoft.SharePoint.Client;
using System.Web;
using System.Linq;
using Claim = System.Security.Claims.Claim;
using ClaimsIdentity = System.Security.Claims.ClaimsIdentity;
using ClaimTypes = System.Security.Claims.ClaimTypes;

namespace AspNet.Owin.SharePoint.Addin.Authentication.Middleware
{
    public class SPAddInAuthenticationHandler : AuthenticationHandler<SPAddInAuthenticationOptions>
    {
        private readonly ILogger _logger;

        public SPAddInAuthenticationHandler(ILogger logger)
        {
            _logger = logger;
        }

        protected override async Task<AuthenticationTicket> AuthenticateCoreAsync()
        {
            Uri spHostUrl;
            Uri spAppWebUrl;
            AccessToken accessToken = null;
            ClaimsIdentity identity = new ClaimsIdentity(Options.SignInAsAuthenticationType);
            identity.AddClaim(new Claim(SPAddinClaimTypes.SPAddinAuthentication, "1"));
            identity.AddClaim(new Claim(SPAddinClaimTypes.ClientId, Options.ClientId.ToString()));

            //if (Context.Authentication.User.Identity.IsAuthenticated &&
            //    Context.Authentication.User.Identity.AuthenticationType == Options.SignInAsAuthenticationType)
            //{
            //    identity = (ClaimsIdentity)Context.Authentication.User.Identity;
            //}
            string spHostUrlString = TokenHelper.EnsureTrailingSlash(Request.Query.Get(SharePointContext.SPHostUrlKey));
            if (string.IsNullOrEmpty(spHostUrlString))
            {
                spHostUrlString = this.Options.SPHostUrl;
            }
            if (!Uri.TryCreate(spHostUrlString, UriKind.Absolute, out spHostUrl))
            {
                throw new Exception($"Unable to determine {SharePointContext.SPHostUrlKey}.");
            }
            else
            {
                identity.AddClaim(new Claim(SPAddinClaimTypes.SPHostUrl, spHostUrl.GetLeftPart(UriPartial.Path)));
            }
            string spAppWebUrlString = TokenHelper.EnsureTrailingSlash(Request.Query.Get(SharePointContext.SPAppWebUrlKey));
            if (!Uri.TryCreate(spAppWebUrlString, UriKind.Absolute, out spAppWebUrl))
            {
                //throw new Exception(string.Format("Unable to determine {0}.", SharePointContext.SPAppWebUrlKey));
            }
            else
            {
                identity.AddClaim(new Claim(SPAddinClaimTypes.SPAppWebUrl, spAppWebUrl.GetLeftPart(UriPartial.Path)));
            }
            //}
            //string accessTokenString = null;
            if (TokenHelper.IsHighTrustApp())
            {
                string userSid = OwinTokenHelper.GetWindowsUserSid(Context);
                //accessTokenString = OwinTokenHelper.GetS2SAccessToken(spHostUrl, userSid);
                identity.AddClaim(new Claim(SPAddinClaimTypes.ADUserId, userSid));
                identity.AddClaim(new Claim(SPAddinClaimTypes.CacheKey, userSid));
                identity.AddClaim(new Claim(SPAddinClaimTypes.Realm, TokenHelper.GetRealmFromTargetUrl(spHostUrl)));

                var spContext = new HighTrustContext(identity);
                try
                {
                    accessToken = spContext.CreateUserAccessTokenForSPHost();
                }
                catch (Microsoft.IdentityModel.SecurityTokenService.RequestFailedException)
                {
                    accessToken = null;
                }
            }
            else
            {
                var contextTokenString = OwinTokenHelper.GetContextTokenFromRequest(Request);
                if (!string.IsNullOrEmpty(contextTokenString))
                {
                    var contextToken = TokenHelper.ReadAndValidateContextToken(contextTokenString, Request.Uri.Authority);
                    if (contextToken != null)
                    {
                        identity.AddClaim(new Claim(SPAddinClaimTypes.RefreshToken, contextToken.RefreshToken));
                        identity.AddClaim(new Claim(SPAddinClaimTypes.Realm, contextToken.Realm));
                        identity.AddClaim(new Claim(SPAddinClaimTypes.TargetPrincipalName, contextToken.TargetPrincipalName));
                        identity.AddClaim(new Claim(SPAddinClaimTypes.CacheKey, contextToken.CacheKey));
                        //OAuth2AccessTokenResponse accessToken = TokenHelper.GetAccessToken(contextToken.RefreshToken, contextToken.TargetPrincipalName, spHostUrl.Authority, contextToken.Realm);
                        //accessTokenString = accessToken.AccessToken;
                        var spContext = new AcsContext(identity);
                        try
                        {
                            accessToken = spContext.CreateUserAccessTokenForSPHost();
                        }
                        catch (Microsoft.IdentityModel.SecurityTokenService.RequestFailedException)
                        {
                            accessToken = null;
                        }
                    }
                }
            }
            return await CreateTicket(accessToken, identity, spHostUrl);
        }

        protected override Task ApplyResponseChallengeAsync()
        {
            if (Response.StatusCode == 401)
            {
                var challenge = Helper.LookupChallenge(Options.AuthenticationType, Options.AuthenticationMode);
                if (challenge == null)
                {
                    return Task.FromResult<object>(null);
                }
                var state = challenge.Properties;
                var hostUrl = new Uri(state.Dictionary[SharePointContext.SPHostUrlKey]);
                state.Dictionary.Remove(SharePointContext.SPHostUrlKey);
                GenerateCorrelationId(state);
                string stateString = Options.StateDataFormat.Protect(state);
                string redirectUri = GetAppContextTokenRequestUrl(hostUrl, stateString);
                _logger.WriteInformation($"Redirecting to SharePoint AppRedirect: {redirectUri}");
                Response.Redirect(redirectUri);
            }
            return Task.FromResult<object>(null);
        }

        private string GetAppContextTokenRequestUrl(Uri hostUrl, string stateString)
        {
            var uriBuilder = new UriBuilder(Request.Uri)
            {
                //Path = /*Options.CallbackPath.Value*/
            };
            var queryNameValueCollection = HttpUtility.ParseQueryString(uriBuilder.Query);
            queryNameValueCollection.Remove(SharePointContext.SPHostUrlKey);
            queryNameValueCollection.Remove(SharePointContext.SPAppWebUrlKey);
            queryNameValueCollection.Remove(SharePointContext.SPLanguageKey);
            queryNameValueCollection.Remove(SharePointContext.SPClientTagKey);
            queryNameValueCollection.Remove(SharePointContext.SPProductNumberKey);
            var redirectUri = $"{uriBuilder.Uri.GetLeftPart(UriPartial.Path)}?{{StandardTokens}}&state={stateString}";
            //var redirectUri = string.Format("{0}?{{StandardTokens}}&{{{2}}}&state={1}", uriBuilder.Uri.GetLeftPart(UriPartial.Path), stateString, SharePointContext.SPAppWebUrlKey);
            var tokenRequestUrl = TokenHelper.GetAppContextTokenRequestUrl(hostUrl.GetLeftPart(UriPartial.Path), WebUtility.UrlEncode(redirectUri));
            return tokenRequestUrl;
        }

        public override async Task<bool> InvokeAsync()
        {
            if (Response.StatusCode == 401)
            {
                return false;
            }
            //if (Options.CallbackPath.HasValue && Options.CallbackPath == Request.Path)
            //{
            _logger.WriteInformation("Receiving contextual information");

            if (TokenHelper.IsHighTrustApp())
            {
                var logonUserIdentity = OwinTokenHelper.GetHttpRequestIdentity(Context);
                // If not authenticated and we are using integrated windows auth, then force user to login
                if (!logonUserIdentity.IsAuthenticated && logonUserIdentity is WindowsIdentity)
                {
                    Response.StatusCode = 418;
                    // Prevent further processing by the owin pipeline.
                    return true;
                }
            }

            var state = Request.Query["state"];
            var properties = Options.StateDataFormat.Unprotect(state);
            if (properties != null && !ValidateCorrelationId(properties, _logger))
            {
                throw new Exception("Correlation failed.");
            }
            var ticket = await AuthenticateAsync();
            if (ticket != null)
            {
                if (Request.User.Identity.IsAuthenticated)
                {
                    //var identity = Request.User.Identity as ClaimsIdentity;
                    //if (identity != null)
                    //{
                    //    var spContext = SPContextProvider.Get(identity);
                    //}
                }

                if (Response.StatusCode != 401)
                {
                    Context.Authentication.SignIn(ticket.Properties, ticket.Identity);
                    if (string.IsNullOrEmpty(ticket.Properties.RedirectUri))
                    {
                        ticket.Properties.RedirectUri = "/";
                    }
                    var urlParsed = ticket.Properties.RedirectUri.Split('?');
                    string url = urlParsed.FirstOrDefault();
                    string queryString = urlParsed.Skip(1).LastOrDefault();
                    var query = HttpUtility.ParseQueryString(queryString ?? "");
                    if (ticket.Properties.Dictionary.ContainsKey(SharePointContext.SPHostUrlKey))
                    {
                        query[SharePointContext.SPHostUrlKey] =
                            ticket.Properties.Dictionary[SharePointContext.SPHostUrlKey];
                    }
                    if (ticket.Properties.Dictionary.ContainsKey(SharePointContext.SPAppWebUrlKey))
                    {
                        query[SharePointContext.SPAppWebUrlKey] =
                            ticket.Properties.Dictionary[SharePointContext.SPAppWebUrlKey];
                    }
                    Response.Redirect($"{url}?{query}");
                }
                // Prevent further processing by the owin pipeline.
                return true;
            }
            //}
            // Let the rest of the pipeline run.
            return false;
        }

        private async Task<AuthenticationTicket> CreateTicket(AccessToken accessToken, ClaimsIdentity identity, Uri spHostUrl)
        {
            if (accessToken == null || !accessToken.IsValid())
            {
                return null;
            }
            var properties = Options.StateDataFormat.Unprotect(Request.Query["state"]) ?? new AuthenticationProperties();
            if (spHostUrl != null)
            {
                using (var clientContext = TokenHelper.GetClientContextWithAccessToken(spHostUrl.GetLeftPart(UriPartial.Path), accessToken.Value))
                {
                    var user = clientContext.Web.CurrentUser;
                    clientContext.Load(user);
                    try
                    {
                        clientContext.ExecuteQuery();
                        identity.AddClaim(new Claim(ClaimTypes.NameIdentifier, user.LoginName, null, Options.AuthenticationType));
                        identity.AddClaim(new Claim(ClaimTypes.Name, user.Title));
                        identity.AddClaim(new Claim(ClaimTypes.Email, user.Email));
                        identity.AddClaim(new Claim(ClaimTypes.Sid, user.Id.ToString()));
                        await Options.Provider.Authenticated(new SPAddinAuthenticatedContext(Context, user, identity));
                        return new AuthenticationTicket(identity, properties);
                    }
                    catch (ServerUnauthorizedAccessException)
                    {
                        return new AuthenticationTicket(identity, properties);
                    }
                }
            }
            return null;
        }
    }
}
