using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExportExcel.Interfaces
{
    /// <summary>
    /// Interface for caching operations
    /// </summary>
    public interface ICacheManager
    {
        /// <summary>
        /// Gets cached value by key
        /// </summary>
        /// <typeparam name="T">Type of cached value</typeparam>
        /// <param name="key">Cache key</param>
        /// <returns>Cached value or default</returns>
        T Get<T>(string key);

        /// <summary>
        /// Sets cached value
        /// </summary>
        /// <typeparam name="T">Type of value to cache</typeparam>
        /// <param name="key">Cache key</param>
        /// <param name="value">Value to cache</param>
        /// <param name="expiry">Optional expiry time</param>
        void Set<T>(string key, T value, TimeSpan? expiry = null);

        /// <summary>
        /// Removes cached value
        /// </summary>
        /// <param name="key">Cache key</param>
        void Remove(string key);

        /// <summary>
        /// Clears all cached values
        /// </summary>
        void Clear();

        /// <summary>
        /// Checks if key exists in cache
        /// </summary>
        /// <param name="key">Cache key</param>
        /// <returns>True if key exists</returns>
        bool Contains(string key);
    }

}
