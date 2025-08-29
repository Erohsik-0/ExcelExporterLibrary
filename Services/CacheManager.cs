using ExportExcel.Interfaces;
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace ExportExcel.Services
{
    /// <summary>
    /// Thread-safe in-memory cache manager for Excel operations
    /// </summary>
    public class CacheManager : ICacheManager, IDisposable
    {
        private readonly ConcurrentDictionary<string, CacheEntry> _cache;
        private readonly Timer _cleanupTimer;
        private readonly TimeSpan _defaultExpiry;
        private readonly int _maxCacheSize;
        private readonly object _lockObject = new object();
        private bool _disposed = false;

        public CacheManager(TimeSpan? defaultExpiry = null, int maxCacheSize = 1000)
        {
            _cache = new ConcurrentDictionary<string, CacheEntry>();
            _defaultExpiry = defaultExpiry ?? TimeSpan.FromMinutes(30);
            _maxCacheSize = maxCacheSize;

            // Setup cleanup timer to run every 5 minutes
            _cleanupTimer = new Timer(CleanupExpiredEntries, null, TimeSpan.FromMinutes(5), TimeSpan.FromMinutes(5));
        }

        public T Get<T>(string key)
        {
            if (string.IsNullOrWhiteSpace(key))
                return default(T);

            if (_cache.TryGetValue(key, out var entry))
            {
                if (!entry.IsExpired)
                {
                    entry.LastAccessed = DateTime.UtcNow;
                    entry.AccessCount++;

                    try
                    {
                        return (T)entry.Value;
                    }
                    catch (InvalidCastException)
                    {
                        // Type mismatch, remove the entry
                        _cache.TryRemove(key, out _);
                        return default(T);
                    }
                }
                else
                {
                    // Entry expired, remove it
                    _cache.TryRemove(key, out _);
                }
            }

            return default(T);
        }

        public void Set<T>(string key, T value, TimeSpan? expiry = null)
        {
            if (string.IsNullOrWhiteSpace(key))
                return;

            var expiryTime = expiry ?? _defaultExpiry;
            var entry = new CacheEntry
            {
                Value = value,
                CreatedAt = DateTime.UtcNow,
                LastAccessed = DateTime.UtcNow,
                ExpiresAt = DateTime.UtcNow.Add(expiryTime),
                AccessCount = 0
            };

            // Check cache size and evict if necessary
            if (_cache.Count >= _maxCacheSize)
            {
                EvictLeastRecentlyUsed();
            }

            _cache.AddOrUpdate(key, entry, (k, oldEntry) => entry);
        }

        public void Remove(string key)
        {
            if (!string.IsNullOrWhiteSpace(key))
            {
                _cache.TryRemove(key, out _);
            }
        }

        public void Clear()
        {
            _cache.Clear();
        }

        public bool Contains(string key)
        {
            if (string.IsNullOrWhiteSpace(key))
                return false;

            if (_cache.TryGetValue(key, out var entry))
            {
                if (!entry.IsExpired)
                {
                    return true;
                }
                else
                {
                    // Entry expired, remove it
                    _cache.TryRemove(key, out _);
                    return false;
                }
            }

            return false;
        }

        #region Advanced Cache Operations

        /// <summary>
        /// Gets cache statistics
        /// </summary>
        public CacheStatistics GetStatistics()
        {
            var entries = _cache.Values.ToList();

            return new CacheStatistics
            {
                TotalEntries = entries.Count,
                ExpiredEntries = entries.Count(e => e.IsExpired),
                TotalHits = entries.Sum(e => e.AccessCount),
                AverageAge = entries.Any() ? TimeSpan.FromTicks((long)entries.Average(e => (DateTime.UtcNow - e.CreatedAt).Ticks)) : TimeSpan.Zero,
                OldestEntry = entries.Any() ? entries.Min(e => e.CreatedAt) : (DateTime?)null,
                NewestEntry = entries.Any() ? entries.Max(e => e.CreatedAt) : (DateTime?)null
            };
        }

        /// <summary>
        /// Gets all cache keys
        /// </summary>
        public IEnumerable<string> GetKeys()
        {
            return _cache.Keys.ToList();
        }

        /// <summary>
        /// Gets or sets a cached value with lazy initialization
        /// </summary>
        public T GetOrSet<T>(string key, Func<T> factory, TimeSpan? expiry = null)
        {
            if (Contains(key))
            {
                return Get<T>(key);
            }

            var value = factory();
            Set(key, value, expiry);
            return value;
        }

        /// <summary>
        /// Async version of GetOrSet
        /// </summary>
        public async Task<T> GetOrSetAsync<T>(string key, Func<Task<T>> factory, TimeSpan? expiry = null)
        {
            if (Contains(key))
            {
                return Get<T>(key);
            }

            var value = await factory();
            Set(key, value, expiry);
            return value;
        }

        /// <summary>
        /// Sets multiple values at once
        /// </summary>
        public void SetMany<T>(Dictionary<string, T> items, TimeSpan? expiry = null)
        {
            if (items == null) return;

            foreach (var item in items)
            {
                Set(item.Key, item.Value, expiry);
            }
        }

        /// <summary>
        /// Gets multiple values at once
        /// </summary>
        public Dictionary<string, T> GetMany<T>(IEnumerable<string> keys)
        {
            var result = new Dictionary<string, T>();

            if (keys == null) return result;

            foreach (var key in keys)
            {
                var value = Get<T>(key);
                if (!EqualityComparer<T>.Default.Equals(value, default(T)))
                {
                    result[key] = value;
                }
            }

            return result;
        }

        /// <summary>
        /// Removes entries matching a pattern
        /// </summary>
        public int RemoveByPattern(string pattern)
        {
            if (string.IsNullOrWhiteSpace(pattern))
                return 0;

            var keysToRemove = _cache.Keys
                .Where(key => key.Contains(pattern, StringComparison.OrdinalIgnoreCase))
                .ToList();

            int removedCount = 0;
            foreach (var key in keysToRemove)
            {
                if (_cache.TryRemove(key, out _))
                {
                    removedCount++;
                }
            }

            return removedCount;
        }

        /// <summary>
        /// Extends expiry time for an existing entry
        /// </summary>
        public bool ExtendExpiry(string key, TimeSpan additionalTime)
        {
            if (_cache.TryGetValue(key, out var entry) && !entry.IsExpired)
            {
                entry.ExpiresAt = entry.ExpiresAt.Add(additionalTime);
                return true;
            }

            return false;
        }

        #endregion

        #region Private Methods

        private void CleanupExpiredEntries(object state)
        {
            if (_disposed) return;

            var expiredKeys = _cache
                .Where(kvp => kvp.Value.IsExpired)
                .Select(kvp => kvp.Key)
                .ToList();

            foreach (var key in expiredKeys)
            {
                _cache.TryRemove(key, out _);
            }
        }

        private void EvictLeastRecentlyUsed()
        {
            lock (_lockObject)
            {
                if (_cache.Count < _maxCacheSize) return;

                // Remove 20% of the least recently used entries
                var entriesToRemove = (int)Math.Ceiling(_cache.Count * 0.2);
                var lruEntries = _cache
                    .OrderBy(kvp => kvp.Value.LastAccessed)
                    .Take(entriesToRemove)
                    .Select(kvp => kvp.Key)
                    .ToList();

                foreach (var key in lruEntries)
                {
                    _cache.TryRemove(key, out _);
                }
            }
        }

        #endregion

        #region IDisposable Implementation

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        protected virtual void Dispose(bool disposing)
        {
            if (!_disposed)
            {
                if (disposing)
                {
                    _cleanupTimer?.Dispose();
                    _cache?.Clear();
                }

                _disposed = true;
            }
        }

        ~CacheManager()
        {
            Dispose(false);
        }

        #endregion

        #region Helper Classes

        private class CacheEntry
        {
            public object Value { get; set; }
            public DateTime CreatedAt { get; set; }
            public DateTime LastAccessed { get; set; }
            public DateTime ExpiresAt { get; set; }
            public long AccessCount { get; set; }

            public bool IsExpired => DateTime.UtcNow > ExpiresAt;
        }

        public class CacheStatistics
        {
            public int TotalEntries { get; set; }
            public int ExpiredEntries { get; set; }
            public long TotalHits { get; set; }
            public TimeSpan AverageAge { get; set; }
            public DateTime? OldestEntry { get; set; }
            public DateTime? NewestEntry { get; set; }
            public double HitRatio => TotalEntries > 0 ? (double)TotalHits / TotalEntries : 0;
        }

        #endregion
    }

}
