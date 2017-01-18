using System;
using System.Collections;
using System.Collections.Generic;

namespace AspNet.Owin.SharePoint.Addin.Authentication.Caching
{
    public class TokenCache : ITokenCache
    {
        protected static CachingProvider _сachingProvider;

        static TokenCache()
        {
            _сachingProvider = new CachingProvider();
        }

        public void Insert(AccessToken token, string key)
        {
            _сachingProvider.AddItem(key, token);
        }

        public void Remove(string key)
        {
            _сachingProvider.RemoveItem(key);
        }

        public AccessToken Get(string key)
        {
            var token = (AccessToken)_сachingProvider.GetItem(key);
            if (token != null && IsAccessTokenValid(token))
            {
                return token;
            }
            _сachingProvider.RemoveItem(key);
            return null;
        }

        public bool IsAccessTokenValid(AccessToken token)
        {
            return !string.IsNullOrEmpty(token?.Value) && token.ExpiresOn > DateTime.UtcNow;
        }
    }

    public class DefaultTokenCache : ITokenCache
    {
        protected static Dictionary<string, AccessToken> _tokens;

        static DefaultTokenCache()
        {
            _tokens = new Dictionary<string, AccessToken>();
        }

        public void Insert(AccessToken token, string key)
        {
            lock (((IDictionary)_tokens).SyncRoot)
            {
                _tokens[key] = token;
            }
        }

        public void Remove(string key)
        {
            lock (((IDictionary)_tokens).SyncRoot)
            {
                if (_tokens.ContainsKey(key))
                {
                    _tokens.Remove(key);
                }
            }
        }

        public AccessToken Get(string key)
        {
            lock (((IDictionary)_tokens).SyncRoot)
            {
                if (_tokens.ContainsKey(key))
                {
                    var token = _tokens[key];
                    if (IsAccessTokenValid(token))
                    {
                        return token;
                    }
                    else
                    {
                        _tokens.Remove(key);
                    }
                }
                return null;
            }
        }

        public bool IsAccessTokenValid(AccessToken token)
        {
            return !string.IsNullOrEmpty(token?.Value) && token.ExpiresOn > DateTime.UtcNow;
        }
    }

}
