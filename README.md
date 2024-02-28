# msgext-graph-search
msgext-graph-search has been built using search-based [message extension](https://docs.microsoft.com/microsoftteams/platform/messaging-extensions/what-are-messaging-extensions?tabs=nodejs) template that allows users to leverages the Graph Search API to search Microsoft 365 data and share results through the compose message area of the Microsoft Teams client. You can now build and run the this message extensions in Teams, Outlook for Windows desktop and web experiences.

https://github.com/mohammadamer/msgext-graph-search/assets/19314043/eb379605-0190-4191-a52e-8a79bb96b721

## Main functionalities
- Login and consent in Message Extension
- Use Microsoft Graph API to query Microsoft 365 data with SSO token in Message Extension

## Prerequisite to use this sample
- [Node.js](https://nodejs.org/), supported versions: 16, 18
- A Microsoft 365 account. If you do not have Microsoft 365 account, apply one from [Microsoft 365 developer program](https://developer.microsoft.com/en-us/microsoft-365/dev-program)
- [Teams Toolkit Visual Studio Code Extension](https://aka.ms/teams-toolkit) version 5.0.0 and higher or [TeamsFx CLI](https://aka.ms/teamsfx-cli)

## Minimal path to awesome
> Here are the instructions to run the sample in **Visual Studio Code**.
### Run the app locally
1. Clone the repo to your local workspace or directly download the source code.
3. Open the project in Visual Studio Code.
4. Open Debug View (`Ctrl+Shift+D`) and select `Debug(Edge)` or `Debug(Chrome)` in dropdown list.
5. Press `F5` finally pop up an window to login your tenant account, and install this Teams App.

### Deploy the app to Azure
Deploy your project to Azure by following these steps:
1. Open Teams Toolkit in Visual Studio Code, and sign in your Azure account by clicking the `Sign in to Azure` in the `ACCOUNTS` section from sidebar.
2. Open Teams Toolkit and click `Provision` in the `LIFECYCLE` section. Alternatively open the command palette(Ctrl+Shift+P) and type: `Teams: Provision` command.
3. Open the Teams Toolkit and click `Deploy` in the `LIFECYCLE` section. Alternatively open the command palette(Ctrl+Shift+P) and type: `Teams: Deploy` command.
> Note: Provision Azure cloud resources and deploy to Azure may cause charges to your Azure Subscription.
### Preview the app in Teams
After you have completed the provision and deploy steps in `Deploy the app to Azure` section, you can preview your app in Teams client by following steps below:

1. Open the `Run and Debug Activity` panel from sidebar, or use short key Ctrl+Shift+D.
1. Select `Launch Remote (Edge)` or `Launch Remote (Chrome)` in the launch configuration (a dropdown selection in the upper-left corner).
1. Press the `Start Debugging` button to launch your app, the Teams web client will be automatically opened in your browser, where you will see your app running remotely from Azure.


## What's included in the template

| Folder       | Contents                                            |
| - | - |
| `.vscode/`    | VSCode files for debugging                          |
| `appPackage/` | Templates for the Teams application manifest        |
| `env/`        | Environment files                                   |
| `infra/`      | Templates for provisioning Azure resources          |
| `src/` | The source code for the search application |

The following files can be customized and demonstrate an example implementation to get you started.

| File                                 | Contents                                           |
| - | - |
|`src/searchApp.ts`| Handles the business logic for this app template to query npm registry and return result list.|
|`src/index.ts`| `index.ts` is used to setup and configure the Message Extension.|

The following are Teams Toolkit specific project files. You can [visit a complete guide on Github](https://github.com/OfficeDev/TeamsFx/wiki/Teams-Toolkit-Visual-Studio-Code-v5-Guide#overview) to understand how Teams Toolkit works.

| File                                 | Contents                                           |
| - | - |
|`teamsapp.yml`|This is the main Teams Toolkit project file. The project file defines two primary things:  Properties and configuration Stage definitions. |
|`teamsapp.local.yml`|This overrides `teamsapp.yml` with actions that enable local execution and debugging.|

## Extend the template

Following documentation will help you to extend the template.

- [Add or manage the environment](https://learn.microsoft.com/microsoftteams/platform/toolkit/teamsfx-multi-env)
- [Create multi-capability app](https://learn.microsoft.com/microsoftteams/platform/toolkit/add-capability)
- [Add single sign on to your app](https://learn.microsoft.com/microsoftteams/platform/toolkit/add-single-sign-on)
- [Access data in Microsoft Graph](https://learn.microsoft.com/microsoftteams/platform/toolkit/teamsfx-sdk#microsoft-graph-scenarios)
- [Use an existing Azure Active Directory application](https://learn.microsoft.com/microsoftteams/platform/toolkit/use-existing-aad-app)
- [Customize the Teams app manifest](https://learn.microsoft.com/microsoftteams/platform/toolkit/teamsfx-preview-and-customize-app-manifest)
- Host your app in Azure by [provision cloud resources](https://learn.microsoft.com/microsoftteams/platform/toolkit/provision) and [deploy the code to cloud](https://learn.microsoft.com/microsoftteams/platform/toolkit/deploy)
- [Collaborate on app development](https://learn.microsoft.com/microsoftteams/platform/toolkit/teamsfx-collaboration)
- [Set up the CI/CD pipeline](https://learn.microsoft.com/microsoftteams/platform/toolkit/use-cicd-template)
- [Publish the app to your organization or the Microsoft Teams app store](https://learn.microsoft.com/microsoftteams/platform/toolkit/publish)
- [Develop with Teams Toolkit CLI](https://aka.ms/teamsfx-cli/debug)
- [Preview the app on mobile clients](https://github.com/OfficeDev/TeamsFx/wiki/Run-and-debug-your-Teams-application-on-iOS-or-Android-client)
