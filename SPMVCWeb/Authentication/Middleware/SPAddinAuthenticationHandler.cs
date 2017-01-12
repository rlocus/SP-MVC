using System;
using System.Net;
using System.Security.Claims;
using System.Security.Principal;
using System.Threading.Tasks;
using AspNet.Owin.SharePoint.Addin.Authentication.Provider;
using Microsoft.Owin.Logging;
using Microsoft.Owin.Security;
using Microsoft.Owin.Security.Infrastructure;
using Microsoft.SharePoint.Client;
using System.Web;
using System.Linq;

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
            ClaimsIdentity identity = new ClaimsIdentity(Options.SignInAsAuthenticationType);

            //if (Context.Authentication.User.Identity.IsAuthenticated &&
            //    Context.Authentication.User.Identity.AuthenticationType == Options.SignInAsAuthenticationType)
            //{

            //}

            string spHostUrlString = TokenHelper.EnsureTrailingSlash(Request.Query.Get(SharePointContext.SPHostUrlKey));
            if (string.IsNullOrEmpty(spHostUrlString))
            {
                spHostUrlString = this.Options.SPHostUrl;
            }
            if (!Uri.TryCreate(spHostUrlString, UriKind.Absolute, out spHostUrl))
            {
                throw new Exception(string.Format("Unable to determine {0}.", SharePointContext.SPHostUrlKey));
            }
            else
            {
                identity.AddClaim(new Claim(SPAddinClaimTypes.SPHostUrl, spHostUrl.AbsoluteUri));
            }
            string spAppWebUrlString = TokenHelper.EnsureTrailingSlash(Request.Query.Get(SharePointContext.SPAppWebUrlKey));
            if (!Uri.TryCreate(spAppWebUrlString, UriKind.Absolute, out spAppWebUrl))
            {
                //throw new Exception(string.Format("Unable to determine {0}.", SharePointContext.SPAppWebUrlKey));
            }
            else
            {
                identity.AddClaim(new Claim(SPAddinClaimTypes.SPAppWebUrl, spAppWebUrl.AbsoluteUri));
            }
            //}
            string accessToken = null;
            if (TokenHelper.IsHighTrustApp())
            {
                var userSid = AuthHelper.GetWindowsUserSid(Context);
                accessToken = AuthHelper.GetS2SAccessToken(spHostUrl, userSid);
                identity.AddClaim(new Claim(SPAddinClaimTypes.ADUserId, userSid));
                identity.AddClaim(new Claim(SPAddinClaimTypes.CacheKey, userSid));
                identity.AddClaim(new Claim(SPAddinClaimTypes.Realm, TokenHelper.GetRealmFromTargetUrl(spHostUrl)));
            }
            else
            {
                var contextTokenString = AuthHelper.GetContextTokenFromRequest(Request);
                if (!string.IsNullOrEmpty(contextTokenString) && spHostUrl != null)
                {
                    var contextToken = TokenHelper.ReadAndValidateContextToken(contextTokenString, Request.Uri.Authority);
                    if (contextToken != null)
                    {
                        identity.AddClaim(new Claim(SPAddinClaimTypes.RefreshToken, contextToken.RefreshToken));
                        identity.AddClaim(new Claim(SPAddinClaimTypes.Realm, contextToken.Realm));
                        identity.AddClaim(new Claim(SPAddinClaimTypes.TargetPrincipalName, contextToken.TargetPrincipalName));
                        identity.AddClaim(new Claim(SPAddinClaimTypes.CacheKey, contextToken.CacheKey));
                        accessToken = TokenHelper.GetAccessToken(contextToken.RefreshToken, contextToken.TargetPrincipalName, spHostUrl.Authority, contextToken.Realm).AccessToken;
                    }
                }
            }
            return await CreateTicket(accessToken, identity, spHostUrl, spAppWebUrl);
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
                _logger.WriteInformation(string.Format("Redirecting to SharePoint AppRedirect: {0}", redirectUri));
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
            var redirectUri = string.Format("{0}?{{StandardTokens}}&state={1}", uriBuilder.Uri.GetLeftPart(UriPartial.Path), stateString);
            //var redirectUri = string.Format("{0}?{{StandardTokens}}&{{{2}}}&state={1}", uriBuilder.Uri.GetLeftPart(UriPartial.Path), stateString, SharePointContext.SPAppWebUrlKey);
            var tokenRequestUrl = TokenHelper.GetAppContextTokenRequestUrl(hostUrl.GetLeftPart(UriPartial.Path), WebUtility.UrlEncode(redirectUri));
            return tokenRequestUrl;
        }

        public override async Task<bool> InvokeAsync()
        {
            //if (Options.CallbackPath.HasValue && Options.CallbackPath == Request.Path)
            //{
            _logger.WriteInformation("Receiving contextual information");

            if (TokenHelper.IsHighTrustApp())
            {
                var logonUserIdentity = AuthHelper.GetHttpRequestIdentity(Context);
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
                if (Response.StatusCode != 401)
                {
                    ticket.Identity.AddClaim(new Claim(SPAddinClaimTypes.SPAddinAuthentication, "1"));
                    ticket.Identity.AddClaim(new Claim(SPAddinClaimTypes.ClientId, Options.ClientId.ToString()));
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
                    Response.Redirect(string.Format("{0}?{1}", url, query));
                }
                // Prevent further processing by the owin pipeline.
                return true;
            }
            //}
            // Let the rest of the pipeline run.
            return false;
        }

        private async Task<AuthenticationTicket> CreateTicket(string accessToken, ClaimsIdentity identity, Uri spHostUrl, Uri spAppWebUrl)
        {
            if (string.IsNullOrEmpty(accessToken))
            {
                return null;
            }
            var properties = Options.StateDataFormat.Unprotect(Request.Query["state"]) ?? new AuthenticationProperties();
            if (spHostUrl != null)
            {
                properties.Dictionary.Add(SharePointContext.SPHostUrlKey, spHostUrl.AbsoluteUri);
                if (spAppWebUrl != null)
                {
                    properties.Dictionary.Add(SharePointContext.SPAppWebUrlKey, spAppWebUrl.AbsoluteUri);
                }

                using (var clientContext = TokenHelper.GetClientContextWithAccessToken(spHostUrl.AbsoluteUri, accessToken))
                {
                    var user = clientContext.Web.CurrentUser;
                    clientContext.Load(user);
                    try
                    {
                        clientContext.ExecuteQuery();
                        identity.AddClaim(new Claim(ClaimTypes.NameIdentifier, user.LoginName, null, Options.AuthenticationType));
                        identity.AddClaim(new Claim(ClaimTypes.Name, user.Title));
                        identity.AddClaim(new Claim(ClaimTypes.Email, user.Email));
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
