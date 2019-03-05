# Completed module: Add Azure AD authentication

The version of the project in this directory reflects completing the tutorial up through [Add Azure AD authentication](https://docs.microsoft.com/graph/training/aspnet-tutorial?tutorial-step=3). If you use this version of the project, you need to complete the rest of the tutorial starting at [Get calendar data](https://docs.microsoft.com/graph/training/aspnet-tutorial?tutorial-step=4).

Installing NuGet packages adds content to the /packages, /Content & /Scripts folders. The files within the /Content & /Scripts folders frequently contain 3rd party libraries which do not include redistributable licenses. Therefore the files in these folders are not included in the repository. The NuGet restore process will not repopulate these folders, only a NuGet reinstall of the package will do that. Therefore you will need to reinstall all packages using the NuGet CLI to repopulate these folders if you want to run the final built solution locally. Refer to the following for more information: [Microsoft Docs – How to reinstall and update packages](https://docs.microsoft.com/en-us/nuget/consume-packages/reinstalling-and-updating-packages).”

> **Note:** It is assumed that you have already registered an application in the app registration portal as specified in [Register the app in the portal](https://docs.microsoft.com/graph/training/aspnet-tutorial?tutorial-step=2). You need to configure this version of the sample as follows:
>
> 1. Rename the `PrivateSettings.config.example` file to `PrivateSettings.config`.
> 1. Edit the `PrivateSettings.config` file and make the following changes.
>     1. Replace `YOUR APP ID HERE` with the **Application Id** you got from the App Registration Portal.
>     1. Replace `YOUR APP PASSWORD HERE` with the password you got from the App Registration Portal.