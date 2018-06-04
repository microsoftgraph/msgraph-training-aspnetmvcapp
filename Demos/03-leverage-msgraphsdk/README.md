# Update the ASP.NET MVC Application to Leverage the Microsoft Graph .NET SDK

In this demo you will add a controller and views that utilize the Microsoft Graph .NET SDK to show the user's calendar.

1. Add the Microsoft Graph .NET SDK via NuGet.

    1. In Visual Studio, select the menu item **View > Other Windows > Package Manager Console**.
    1. In the **Package Manager Console** tool window, run the following command to install the Microsoft Graph .NET SDK:

        ```powershell
        Install-Package Microsoft.Graph
        ```

1. Create a model class to store the event information obtained from the Microsoft Graph API:

    1. In the **Visual Studio** **Solution Explorer** tool window, right-click the **Models** folder and select **Add > Class**.
    1. In the **Add Class** dialog, name the class **MyEvent** and select **Add**.
    1. Add the following `using` statements to the existing ones in the **MyEvents.cs** file that was created.

        ```cs
        using System.ComponentModel;
        using System.ComponentModel.DataAnnotations;
        ```

    1. Add the following code to the `MyEvent` class defining three new members:

        ```cs
        [DisplayName("Subject")]
        public string Subject { get; set; }

        [DisplayName("Start Time")]
        [DisplayFormat(DataFormatString = "{0:MM/dd/yyyy}", ApplyFormatInEditMode = true)]
        public DateTimeOffset? Start { get; set; }

        [DisplayName("End Time")]
        [DisplayFormat(DataFormatString = "{0:MM/dd/yyyy}", ApplyFormatInEditMode = true)]
        public DateTimeOffset? End { get; set; }
        ```

1. Update the authentication provider by adding methods that will return the Microsoft Graph SDK client object:

    1. Open the **Helpers/SampleAuthProvider.cs** file.
    1. Add the following `using` statements to the top of the file:

        ```cs
        using System.Net.Http.Headers;
        using Microsoft.Graph;
        ```

    1. Add the following members to the `SampleAuthProvider` class to return the object and handle signing out:

        ```cs
        private static GraphServiceClient graphClient = null;

        public static GraphServiceClient GetAuthenticatedClient()
        {
          GraphServiceClient graphClient = new GraphServiceClient(
            new DelegateAuthenticationProvider(
              async (requestMessage) =>
              {
                string accessToken = await SampleAuthProvider.Instance.GetUserAccessTokenAsync();
                requestMessage.Headers.Authorization = new AuthenticationHeaderValue("bearer", accessToken);
              }));
          return graphClient;
        }

        public static void SignOutClient()
        {
          graphClient = null;
        }
        ```

1. Update the application to handle Microsoft Graph specific exceptions and signing out of the application:

    1. Open the **Helpers/SampleAuthProvider.cs** file.
    1. Locate the `try-catch` statement at the end of the `GetUserAccessTokenAsync()` method. Replace the following line that throws a generic exception...

        ```cs
        throw new Exception ();
        ```

        With the following code to throw a Microsoft Graph specific exception:

        ```cs
        throw new ServiceException(
          new Error
          {
            Code = GraphErrorCode.AuthenticationFailure.ToString(),
            Message = Resource.Error_AuthChallengeNeeded,
          });
        ```

    1. Open the **Controllers/AccountController.cs** file.
    1. Add the following `using` statements to the top of the file:

        ```cs
        using System.Net.Http.Headers;
        using Microsoft.Graph;
        ```

    1. Locate the `SignoOut()` method. Add the following statement immediately before the `Response.Redirect("/");` line:

      ```cs
      SampleAuthProvider.SignOutClient();
      ```

1. Add a new ASP.NET MVC controller that will retrieve events from the user's calendar:

    1. In the **Visual Studio** **Solution Explorer** tool window, right-click the **Controllers** folder and select **Add > Controller**.
    1. In the **Add Scaffold** dialog, select **MVC 5 Controller - Empty**, select **Add** and name the controller **CalendarController** and then select **Add**.
    1. Add the following `using` statements to the existing ones in the **CalendarController.cs** file that was created.

        ```cs
        using Microsoft.Graph;
        using MSGraphCalendarViewer.Helpers;
        using MSGraphCalendarViewer.Models;
        using System.Net.Http.Headers;
        using System.Security.Claims;
        using System.Threading.Tasks;
        ```

    1. Decorate the controller to allow only authenticated users to use it by adding `[Authorize]` in the line immediately before the controller:

        ```cs
        [Authorize]
        public class CalendarController : Controller
        ```

    1. Modify the existing `Index()` method to be asynchronous by adding the `async` keyword and modifying the return type to be as follows:

        ```cs
        public async Task<ActionResult> Index()
        ```

    1. Update the `Index()` method to use the `GraphServiceClient` object to call the Microsoft Graph API and retrieve the first 20 events in the user's calendar:

        ```cs
        public async Task<ActionResult> Index()
        {
          var eventsResults = new List<MyEvent>();

          try
          {
            var graphService = SDKHelper.GetAuthenticatedClient();
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
        ```

1. Implement the Calendar controller's associated ASP.NET MVC view:

    1. In the `CalendarController` class method `Index()`, locate the `View()` return statement at the end of the method. Right-click `View()` in the code and select **Add View**:

        ![Screenshot adding a view using the context menu in the code.](../../Images/vs-calendarController-01.png)

    1. In the **Add View** dialog, set the following values (*leave all other values as their default values*) and select **Add**:

        * **View name:** Index
        * **Template:** Empty (without model)

    1. In the newly created **Views/Calendar/Index.cshtml** file, replace the default code with the following code:

        ```html
        @{
          ViewBag.Title = "Home Page";
        }
        <div>
          <table>
            <thead>
              <tr>
                <th>Subject</th>
                <th>Start</th>
                <th>End</th>
              </tr>
            </thead>
            <tbody>
              @foreach (var o365Event in ViewBag.Events)
              {
                <tr>
                  <td>@o365Event.Subject</td>
                  <td>@o365Event.Start</td>
                  <td>@o365Event.End</td>
                </tr>
              }
            </tbody>
          </table>
        </div>
        ```

    1. Update the navigation in the **Views/Shared/_Layout.cshtml** file to include a fourth link pointing to a new controller *Calendar*:

        ```html
        <ul class="nav navbar-nav">
          <li>@Html.ActionLink("Home", "Index", "Home")</li>
          <li>@Html.ActionLink("About", "About", "Home")</li>
          <li>@Html.ActionLink("Contact", "Contact", "Home")</li>
          <li>@Html.ActionLink("Calendar", "Index", "Calendar")</li>
        </ul>
        ```

1. Save your changes to all files.

Test the application:

1. Press **F5** to start the application.
1. When the browser loads, select **Signin with Microsoft** and login.
1. If this is the first time running the application, you will be prompted to consent to the application. Review the consent dialog and select **Accept**. The dialog will look similar to the following dialog:

    ![Screesnhot of Azure AD consent dialog](../../Images/aad-consent.png)

1. When the ASP.NET application loads, select the **Calendar** link in the top navigation.
1. You should see a list of calendar items from your calendar appear on the page.

    ![Screesnhot of the web application showing calendar events](../../Images/calendar-events-01.png)
