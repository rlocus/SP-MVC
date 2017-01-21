using System;
using System.Collections;
using System.Collections.Generic;
using System.Runtime.Caching;

namespace AspNet.Owin.SharePoint.Addin.Authentication.Caching
{
    public sealed class DefaultTokenCache : ITokenCache
    {
        private static readonly MemoryCache Cache;
        private static readonly object padLock = new object();
        private const int DEFAULT_CACHE_LIFETIME_MINUTES = 60;

        static DefaultTokenCache()
        {
            Cache = MemoryCache.Default;
        }

        public void Insert(AccessToken token, string key)
        {
            lock (padLock)
            {
                if (Cache.Contains(key))
                {
                    Cache.Remove(key);
                }
                CacheItem item = new CacheItem(key, token);
                CacheItemPolicy policy = new CacheItemPolicy
                {
                    AbsoluteExpiration = DateTimeOffset.Now.AddMinutes(DEFAULT_CACHE_LIFETIME_MINUTES),
                    RemovedCallback = (args) =>
                    {
                        (args.CacheItem.Value as IDisposable)?.Dispose();
                    }
                };
                Cache.Set(item, policy);
            }
        }

        public void Remove(string key)
        {
            lock (padLock)
            {
                var result = Cache[key];
                if (result != null)
                {
                    Cache.Remove(key);
                }
            }
        }

        public AccessToken Get(string key)
        {
            lock (padLock)
            {
                var token = (AccessToken)Cache[key];
                return token;
            }
        }
    }

    public sealed class TokenCache : ITokenCache
    {
        private static readonly Dictionary<string, AccessToken> _tokens;

        static TokenCache()
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
                    return _tokens[key];
                }
                return null;
            }
        }
    }
}
