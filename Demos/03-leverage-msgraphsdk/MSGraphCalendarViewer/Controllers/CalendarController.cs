using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using Microsoft.Graph;
using MSGraphCalendarViewer.Helpers;
using MSGraphCalendarViewer.Models;
using System.Net.Http.Headers;
using System.Security.Claims;
using System.Threading.Tasks;

namespace MSGraphCalendarViewer.Controllers
{
  [Authorize]
  public class CalendarController : Controller
  {
    // GET: Calendar
    public async Task<ActionResult> Index()
    {
      var eventsResults = new List<MyEvent>();

      try
      {
        var graphService = SampleAuthProvider.GetAuthenticatedClient();
        var request = graphService.Me.Events.Request(new Option[] { new QueryOption("top", "20"), new QueryOption("skip", "0") });
        var userEventsCollectionPage = await request.GetAsync();
        foreach (var evnt in userEventsCollectionPage)
        {
          eventsResults.Add(new MyEvent
          {
            Subject = !string.IsNullOrEmpty(evnt.Subject) ? evnt.Subject : string.Empty,
            Start = !string.IsNullOrEmpty(evnt.Start.DateTime) ? DateTime.Parse(evnt.Start.DateTime) : new DateTime(),
            End = !string.IsNullOrEmpty(evnt.End.DateTime) ? DateTime.Parse(evnt.End.DateTime) : new DateTime()

          });
        }
      }
      catch (Exception el)
      {
        el.ToString();
      }

      ViewBag.Events = eventsResults.OrderBy(c => c.Start);

      return View();
    }
  }
}