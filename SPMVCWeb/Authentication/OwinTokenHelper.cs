using System;
using System.Collections.Generic;
using System.IO;
using System.Security.Claims;
using System.Security.Principal;
using System.Text;
using Microsoft.IdentityModel.S2S.Tokens;
using Microsoft.Owin;
using FormCollection = Microsoft.Owin.FormCollection;

namespace AspNet.Owin.SharePoint.Addin.Authentication
{
    internal static class OwinTokenHelper
    {
		public static string GetS2SAccessToken(Uri applicationUri, string userId)
		{
			var realm = TokenHelper.GetRealmFromTargetUrl(applicationUri);
			JsonWebTokenClaim[] claims = null;
			if (userId != null)
			{
				claims = new[]
				{
					new JsonWebTokenClaim(JsonWebTokenConstants.ReservedClaims.NameIdentifier, userId.ToLower()),
					new JsonWebTokenClaim("nii", "urn:office:idp:activedirectory")
				};
			}
			return TokenHelper.GetS2SAccessTokenWithClaims(applicationUri.Authority, realm, claims);
		}

		public static string GetContextTokenFromRequest(IOwinRequest request)
		{
            //return TokenHelper.GetContextTokenFromRequest(((System.Web.HttpContextBase)request.Environment["System.Web.HttpContextBase"]).Request);

            using (var reader = new StreamReader(request.Body, Encoding.UTF8, true))
            {
                var formData = GetForm(reader.ReadToEnd());

                string[] paramNames = { "AppContext", "AppContextToken", "AccessToken", "SPAppToken" };

                foreach (string paramName in paramNames)
                {
                    if (!string.IsNullOrEmpty(formData[paramName]))
                    {
                        return formData[paramName];
                    }
                    if (!string.IsNullOrEmpty(request.Query[paramName]))
                    {
                        return request.Query[paramName];
                    }
                }
            }
            return null;
        }

		public static string GetWindowsUserSid(IOwinContext context)
		{
			if (context.Authentication.User.Identity.IsAuthenticated &&
			    context.Authentication.User.HasClaim(c => c.Type == ClaimTypes.PrimarySid))
			{
				return context.Authentication.User.FindFirst(c => c.Type == ClaimTypes.PrimarySid).Value;
			}

			var httpRequest = ((System.Web.HttpContextBase)context.Environment["System.Web.HttpContextBase"]).Request;
		    return httpRequest.LogonUserIdentity?.FindFirst(c => c.Type == ClaimTypes.PrimarySid).Value;
		}

		public static IIdentity GetHttpRequestIdentity(IOwinContext context)
		{
			var httpRequest = ((System.Web.HttpContextBase)context.Environment["System.Web.HttpContextBase"]).Request;
			return httpRequest.LogonUserIdentity;
		}

		private static readonly Action<string, string, object> AppendItemCallback = (name, value, state) =>
		{
			var dictionary = (IDictionary<string, List<String>>)state;
			List<string> existing;
			if (!dictionary.TryGetValue(name, out existing))
			{
				dictionary.Add(name, new List<string>(1) { value });
			}
			else
			{
				existing.Add(value);
			}
		};

		private static IFormCollection GetForm(string text)
		{
			IDictionary<string, string[]> form = new Dictionary<string, string[]>(StringComparer.OrdinalIgnoreCase);
			var accumulator = new Dictionary<string, List<string>>(StringComparer.OrdinalIgnoreCase);
			ParseDelimited(text, new[] { '&' }, AppendItemCallback, accumulator);
			foreach (var kv in accumulator)
			{
				form.Add(kv.Key, kv.Value.ToArray());
			}
			return new FormCollection(form);
		}

		private static void ParseDelimited(string text, char[] delimiters, Action<string, string, object> callback, object state)
		{
			int textLength = text.Length;
			int equalIndex = text.IndexOf('=');
			if (equalIndex == -1)
			{
				equalIndex = textLength;
			}
			int scanIndex = 0;
			while (scanIndex < textLength)
			{
				int delimiterIndex = text.IndexOfAny(delimiters, scanIndex);
				if (delimiterIndex == -1)
				{
					delimiterIndex = textLength;
				}
				if (equalIndex < delimiterIndex)
				{
					while (scanIndex != equalIndex && char.IsWhiteSpace(text[scanIndex]))
					{
						++scanIndex;
					}
					string name = text.Substring(scanIndex, equalIndex - scanIndex);
					string value = text.Substring(equalIndex + 1, delimiterIndex - equalIndex - 1);
					callback(
						Uri.UnescapeDataString(name.Replace('+', ' ')),
						Uri.UnescapeDataString(value.Replace('+', ' ')),
						state);
					equalIndex = text.IndexOf('=', delimiterIndex);
					if (equalIndex == -1)
					{
						equalIndex = textLength;
					}
				}
				scanIndex = delimiterIndex + 1;
			}
		}
	}
}
