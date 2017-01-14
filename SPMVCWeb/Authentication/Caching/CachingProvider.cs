﻿using System;
using System.Runtime.Caching;

namespace AspNet.Owin.SharePoint.Addin.Authentication.Caching
{
    public interface ICachingProvider
    {
        T GetFromCache<T>(string key, Func<T> cacheMissCallback);        
    }

    public class CachingProvider : ICachingProvider
    {
        protected MemoryCache _cache = MemoryCache.Default;
        protected static readonly object padLock = new object();
        private const int DEFAULT_CACHE_LIFETIME_MINUTES = 60;

        public void AddItem(string key, object value)
        {
            lock (padLock)
            {
                CacheItem item = new CacheItem(key, value);
                CacheItemPolicy policy = new CacheItemPolicy();
                policy.AbsoluteExpiration = DateTimeOffset.Now.AddMinutes(DEFAULT_CACHE_LIFETIME_MINUTES);
                policy.RemovedCallback = new CacheEntryRemovedCallback((args) =>
                {
                    if (args.CacheItem.Value is IDisposable)
                    {
                        ((IDisposable)args.CacheItem.Value).Dispose();
                    }
                });
                _cache.Set(item, policy);
            }
        }

        public object GetItem(string key)
        {
            lock (padLock)
            {
                var result = _cache[key];
                return result;
            }
        }

        public void RemoveItem(string key)
        {
            lock (padLock)
            {
                var result = _cache[key];
                if (result != null)
                {
                    _cache.Remove(key);
                }
            }            
        }

        public T GetFromCache<T>(string key, Func<T> cacheMissCallback)
        {
            var objectFromCache = GetItem(key);
            T objectToReturn = default(T);
            if (objectFromCache == null)
            {
                objectToReturn = cacheMissCallback();
                if (objectToReturn != null)
                {
                    AddItem(key, objectToReturn);
                }
            }
            else
            {
                if (objectFromCache is T)
                {
                    objectToReturn = (T)objectFromCache;
                }
            }
            return objectToReturn;
        }      
    }
}