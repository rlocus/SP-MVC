using System;
using System.Collections;
using System.Collections.Generic;

namespace AspNet.Owin.SharePoint.Addin.Authentication.Caching
{
	public class DefaultTokenCache : ITokenCache
	{
		//protected static Dictionary<string, AccessToken> _tokens;

        protected static CachingProvider _сachingProvider;


        static DefaultTokenCache()
		{
            //_tokens = new Dictionary<string, AccessToken>();
            _сachingProvider = new CachingProvider();
        }

        public void Insert(AccessToken token, string key)
        {
            if (_сachingProvider.GetItem(key) == null)
            {
                _сachingProvider.AddItem(key, token);
            }
            //lock (((IDictionary)_tokens).SyncRoot)
            //{
            //    _tokens[key] = token;
            //}
        }

        public void Remove(string key)
        {
            _сachingProvider.RemoveItem(key);
            //lock (((IDictionary)_tokens).SyncRoot)
            //{
            //    if (_tokens.ContainsKey(key))
            //    {
            //        _tokens.Remove(key);
            //    }
            //}
        }

        public AccessToken Get(string key)
		{
            var token = (AccessToken)_сachingProvider.GetItem(key);
            if (token != null && IsAccessTokenValid(token))
            {
                return token;
            }
            _сachingProvider.RemoveItem(key);
            //if (_tokens.ContainsKey(key) && IsAccessTokenValid(_tokens[key]))
            //{
            //	return _tokens[key];
            //}
            //Remove(key);
            return null;
		}

        public bool IsAccessTokenValid(AccessToken token)
        {
            return !string.IsNullOrEmpty(token?.Value) && token.ExpiredOn > DateTime.UtcNow;
        }
    }
}
