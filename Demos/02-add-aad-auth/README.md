# Completed module: Add Azure AD authentication

The version of the project in this directory reflects completing the tutorial up through [Add Azure AD authentication](https://docs.microsoft.com/graph/tutorials/aspnet?tutorial-step=3). If you use this version of the project, you need to complete the rest of the tutorial starting at [Get calendar data](https://docs.microsoft.com/graph/tutorials/aspnet?tutorial-step=4).

> **Note:** It is assumed that you have already registered an application in the app registration portal as specified in [Register the app in the portal](https://docs.microsoft.com/graph/tutorials/aspnet?tutorial-step=2). You need to configure this version of the sample as follows:
>
> 1. Determine your ASP.NET applications's SSL URL. In Visual Studio's Solution Explorer, select the 
**graph-tutorial** project. In the **Properties** window, find the value of **SSL URL**. Copy this value.
>
>       ![Screenshot of the Visual Studio Properties window](/tutorial/images/vs-project-url.png)
> 1. Rename the `PrivateSettings.config.example` file to `PrivateSettings.config`.
> 1. Edit the `PrivateSettings.config` file and make the following changes.
>     1. Replace `YOUR APP ID HERE` with the **Application Id** you got from the App Registration Portal.
>     1. Replace `YOUR APP PASSWORD HERE` with the password you got from the App Registration Portal.
