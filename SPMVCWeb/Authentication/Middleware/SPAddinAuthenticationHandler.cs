using System;
using System.Net;
using System.Runtime.Remoting.Messaging;
using System.Security.Claims;
using System.Security.Principal;
using System.Threading.Tasks;
using AspNet.Owin.SharePoint.Addin.Authentication.Common;
using AspNet.Owin.SharePoint.Addin.Authentication.Provider;
using Microsoft.Owin.Logging;
using Microsoft.Owin.Security;
using Microsoft.Owin.Security.Infrastructure;
using AspNet.Owin.SharePoint.Addin.Authentication.Context;
using Microsoft.SharePoint.Client;

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
            ClaimsIdentity identity;
            if (Context.Authentication.User.Identity.IsAuthenticated &&
                Context.Authentication.User.Identity.AuthenticationType == Options.SignInAsAuthenticationType)
            {
                identity = (ClaimsIdentity)Context.Authentication.User.Identity;
            }
            else
            {
                identity = new ClaimsIdentity(Options.SignInAsAuthenticationType);
            }
            Uri spHostUrl;
            if (!Uri.TryCreate(Request.Query[SharePointContext.SPHostUrlKey], UriKind.Absolute, out spHostUrl))
            {
                throw new Exception("Cannot get host url from query string");
            }
            Uri spAppWebUrl;
            if (Uri.TryCreate(Request.Query[SharePointContext.SPAppWebUrlKey], UriKind.Absolute, out spAppWebUrl))
            {
                identity.AddClaim(new Claim(SPAddinClaimTypes.SPAppWebUrl, spAppWebUrl.AbsoluteUri));
            }
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
                if (!string.IsNullOrEmpty(contextTokenString))
                {
                    var contextToken = TokenHelper.ReadAndValidateContextToken(contextTokenString, Request.Uri.Authority);
                    if (contextToken != null)
                    {
                        identity.AddClaim(new Claim(SPAddinClaimTypes.RefreshToken, contextToken.RefreshToken));
                        identity.AddClaim(new Claim(SPAddinClaimTypes.Realm, contextToken.Realm));
                        identity.AddClaim(new Claim(SPAddinClaimTypes.TargetPrincipalName, contextToken.TargetPrincipalName));
                        identity.AddClaim(new Claim(SPAddinClaimTypes.CacheKey, contextToken.CacheKey));
                        accessToken = TokenHelper.GetAccessToken(contextToken.RefreshToken, contextToken.TargetPrincipalName,
                                spHostUrl.Authority, contextToken.Realm).AccessToken;
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
                _logger.WriteInformation("Redirecting to SharePoint AppRedirect");
                Response.Redirect(redirectUri);
            }
            return Task.FromResult<object>(null);
        }

        private string GetAppContextTokenRequestUrl(Uri hostUrl, string stateString)
        {
            var uriBuilder = new UriBuilder(Request.Uri)
            {
                Path = Options.CallbackPath.Value
            };
            var postRedirectUrl = string.Format("{0}?{{StandardTokens}}&SPAppWebUrl={{SPAppWebUrl}}&state={1}", uriBuilder.Uri.GetLeftPart(UriPartial.Path), stateString);
            var redirectUri = TokenHelper.GetAppContextTokenRequestUrl(hostUrl.AbsoluteUri, WebUtility.UrlEncode(postRedirectUrl));
            return redirectUri;
        }

        public override async Task<bool> InvokeAsync()
        {
            if (Options.CallbackPath.HasValue && Options.CallbackPath == Request.Path)
            {
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
                        Context.Authentication.SignIn(ticket.Properties, ticket.Identity);
                        Response.Redirect(ticket.Properties.RedirectUri);
                    }
                    // Prevent further processing by the owin pipeline.
                    return true;
                }
            }

            // Let the rest of the pipeline run.
            return false;
        }

        private async Task<AuthenticationTicket> CreateTicket(string accessToken, ClaimsIdentity identity, Uri spHostUrl)
        {
            var properties = Options.StateDataFormat.Unprotect(Request.Query["state"]) ?? new AuthenticationProperties();
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
                    identity.AddClaim(new Claim(SPAddinClaimTypes.SPHostUrl, spHostUrl.AbsoluteUri));
                    await Options.Provider.Authenticated(new SPAddinAuthenticatedContext(Context, user, identity));
                }
                catch (ServerUnauthorizedAccessException)
                {
                    Response.StatusCode = 401;
                    //properties.RedirectUri = GetAppContextTokenRequestUrl(spHostUrl, "");
                    //throw new UnauthorizedAccessException(e.Message);
                }
                catch (Exception)
                {
                    Response.StatusCode = 401;
                    //properties.RedirectUri = GetAppContextTokenRequestUrl(spHostUrl, "");
                }
                return new AuthenticationTicket(identity, properties);
            }
        }
    }
}
