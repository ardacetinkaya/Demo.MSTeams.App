# Demo.MSTeams.App
<img src="https://user-images.githubusercontent.com/4550197/117008179-1daf7300-acf3-11eb-8d1b-d860aeba274a.png" align="left" width="80px" />
 
This repository is a simple demonstration of Microsoft Teams **Tab Application** model with current **ASP.NET Core React App template**. Basically you can see usage of **MS Teams Javascript client SDK** usage and some approaches to use **Microsoft Graph API**. For initial installation of SDK, check [Teams Javascript client SDK](https://docs.microsoft.com/en-us/microsoftteams/platform/tabs/how-to/using-teams-client-sdk)

![image](https://user-images.githubusercontent.com/4550197/117016688-9f0b0380-acfb-11eb-9796-3e00afed968c.png)



When you are ready with development environemnt, first you need to import required **@microsoft/teams-js** package. 

```javascript
import * as microsoftTeams from "@microsoft/teams-js";
```

After initializing the SDK, any SDK methods can be executed in required places in code for example; for a component getting context can be done is componentDidMount()

```javascript
componentDidMount() {
        // .initialize() method is required.
        microsoftTeams.initialize();

        microsoftTeams.getContext((context) => {
            //You can check context's other properties to see what you have
            microsoftTeams.authentication.getAuthToken({
                successCallback: (token) => {
                    microsoftTeams.appInitialization.notifySuccess();
                    //Now you have MS Teams user identification token
                    //You can use this token to identify your self within GraphAPI
                    //When calling an API from GraphAPI you need to create/have an authoraztion token for Graph API
                },
                failureCallback: (error) => {
                    microsoftTeams.appInitialization.notifyFailure({
                        reason: microsoftTeams.appInitialization.FailedReason.AuthFailed,
                        error,
                    });
                }
            });

        });
    }
```


With Microsoft Graph API calls some additional actions can be executed. Tabs application model for Microsoft Teams supports single sign-on (SSO). So logged user in MS Teams can be also a logged-in user for your application. To have this support, you need to define a AAD application in you M365 tenant. **[Create your AAD application](https://docs.microsoft.com/en-us/microsoftteams/platform/tabs/how-to/authentication/auth-aad-sso#1-create-your-aad-application)**

When you have created AAD application, in manifest file additional definition should be done to attach MS Teams app. with AAD application.
```json
"webApplicationInfo": {
  "id": "00000000-0000-0000-0000-000000000000",
  "resource": "api://subdomain.example.com/00000000-0000-0000-0000-000000000000"
}
```


To add an application to Microsoft Teams, a manifest file is required. This manifest file defines some properties for the application. You can check full manifest file description [here](https://docs.microsoft.com/en-us/microsoftteams/platform/resources/schema/manifest-schema).

For this repository check more basic and simple version ---> **[manifest.json](https://github.com/ardacetinkaya/Demo.MSTeams.App/blob/main/Manifest/manifest.json)**

