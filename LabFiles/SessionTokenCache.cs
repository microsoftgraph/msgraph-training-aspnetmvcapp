using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Web;
using Microsoft.Identity.Client;

namespace MSGraphCalendarViewer.TokenStorage
{
  public class SessionTokenCache
  {
    private static ReaderWriterLockSlim SessionLock = new ReaderWriterLockSlim(LockRecursionPolicy.NoRecursion);
    string UserId = string.Empty;
    string CacheId = string.Empty;
    HttpContextBase httpContext = null;

    TokenCache cache = new TokenCache();

    public SessionTokenCache(string userId, HttpContextBase httpcontext)
    {
      UserId = userId;
      CacheId = UserId + "_TokenCache";
      httpContext = httpcontext;
      Load();
    }

    public TokenCache GetMsalCacheInstance()
    {
      cache.SetBeforeAccess(BeforeAccessNotification);
      cache.SetAfterAccess(AfterAccessNotification);
      Load();
      return cache;
    }

    public void SaveUserStateValue(string state)
    {
      SessionLock.EnterWriteLock();
      httpContext.Session[CacheId + "_state"] = state;
      SessionLock.ExitWriteLock();
    }
    public string ReadUserStateValue()
    {
      string state = string.Empty;
      SessionLock.EnterReadLock();
      state = (string)httpContext.Session[CacheId + "_state"];
      SessionLock.ExitReadLock();
      return state;
    }
    public void Load()
    {
      SessionLock.EnterReadLock();
      cache.Deserialize((byte[])httpContext.Session[CacheId]);
      SessionLock.ExitReadLock();
    }

    public void Persist()
    {
      SessionLock.EnterWriteLock();

      cache.HasStateChanged = false;

      httpContext.Session[CacheId] = cache.Serialize();
      SessionLock.ExitWriteLock();
    }

    void BeforeAccessNotification(TokenCacheNotificationArgs args)
    {
      Load();
    }

    void AfterAccessNotification(TokenCacheNotificationArgs args)
    {
      if (cache.HasStateChanged)
      {
        Persist();
      }
    }

  }
}