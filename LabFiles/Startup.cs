using Owin;
using Microsoft.Owin;

[assembly: OwinStartup(typeof(MSGraphCalendarViewer.Startup))]

namespace MSGraphCalendarViewer
{
  public partial class Startup
  {
    public void Configuration(IAppBuilder app)
    {
      ConfigureAuth(app);
    }
  }
}