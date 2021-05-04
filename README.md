# Demo.MSTeams.App
 
Simple demonstration of Microsoft Teams Tab Application model with in ASP.NET Core React App template.


First you need to import required **@microsoft/teams-js** package. Check [Teams Javascript client SDK](https://docs.microsoft.com/en-us/microsoftteams/platform/tabs/how-to/using-teams-client-sdk)

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
