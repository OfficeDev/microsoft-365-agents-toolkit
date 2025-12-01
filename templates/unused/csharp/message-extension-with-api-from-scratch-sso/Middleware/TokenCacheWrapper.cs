using System;
using System.Collections.Concurrent;
using System.Threading.Tasks;

public class TokenCacheWrapper
{
    private readonly ConcurrentDictionary<string, Task<string>> _cache;
    private readonly Func<string, Task<string>> _loadFunc;

    public TokenCacheWrapper(Func<string, Task<string>> loadFunc, int cacheMaxEntries = 5, TimeSpan? cacheMaxAge = null)
    {
        _loadFunc = loadFunc;
        _cache = new ConcurrentDictionary<string, Task<string>>();
    }

    public Task<string> GetSigningKeyAsync(string kid)
    {
        return _cache.GetOrAdd(kid, _loadFunc);
    }

    public void ClearCache()
    {
        _cache.Clear();
    }

    public void RemoveKey(string kid)
    {
        _cache.TryRemove(kid, out _);
    }
}