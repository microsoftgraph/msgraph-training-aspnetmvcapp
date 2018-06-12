using System.Threading.Tasks;

namespace MSGraphCalendarViewer.Helpers
{
  public interface IAuthProvider
  {
    Task<string> GetUserAccessTokenAsync();
  }
}