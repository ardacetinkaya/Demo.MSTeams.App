# Demo.MSTeams.App
<img src="https://user-images.githubusercontent.com/4550197/117008179-1daf7300-acf3-11eb-8d1b-d860aeba274a.png" align="left" width="80px" />
 
This repository is a simple demonstration of Microsoft Teams **Tab Application** model with current **ASP.NET Core React App template**. This repository is kind of a demo to integrate a web application into MS Teams.

Basically you can see usage of **MS Teams Javascript client SDK** usage and some approaches to use **Microsoft Graph API**. For initial installation of SDK, check [Teams Javascript client SDK](https://docs.microsoft.com/en-us/microsoftteams/platform/tabs/how-to/using-teams-client-sdk)


![image](https://user-images.githubusercontent.com/4550197/117025467-99192080-ad03-11eb-9c01-194b536a1cee.png)



With Microsoft Graph API calls some additional actions can be executed. Tabs application model for Microsoft Teams supports single sign-on (SSO). So logged user in MS Teams can be also a logged-in user for your application. To have this support, you need to define a AAD application in you M365 tenant. **[Create your AAD application](https://docs.microsoft.com/en-us/microsoftteams/platform/tabs/how-to/authentication/auth-aad-sso#1-create-your-aad-application)**

When you have created AAD application, in manifest file additional definition should be done to attach MS Teams app. with AAD application.
```json
"webApplicationInfo": {
  "id": "00000000-0000-0000-0000-000000000000",
  "resource": "api://subdomain.example.com/00000000-0000-0000-0000-000000000000"
}
```
When you are ready with development environment, first you need to import required **@microsoft/teams-js** package. 

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
MS Teams' authentication token is not a valid token for Graph API calls. So to make a Graph API calls, first you need the proper token. In this repository with additional back-end API call, MS Teams identification token is used to have proper Graph API token and to have Graph API request.

```javascript
microsoftTeams.getContext((context) => {
            this.setState({ userName: context.userObjectId });
            microsoftTeams.authentication.getAuthToken({
                successCallback: (token) => {
                    microsoftTeams.appInitialization.notifySuccess();
                    //With MS Teams' user token, some back-end API call is done
                    //Within this call, this token is used to have a proper Graph API token
                    //and a GET request is done to the 'graph/me' endpoint.
                    //Then .displayName property is set in component state
                    fetch(`api/graph/beta/me`,
                        {
                            method: "GET",
                            headers: {
                                "Authorization": `${token}`
                            }
                        })
                        .then(response => response.json())
                        .then(data => {
                            if (data.error) {
                                console.error("!!!ERROR!!!");
                                console.error(data.error);
                            }
                            else {
                                console.log(data);
                                this.setState({ userName: data.data.displayName });
                            }

                        })
                        .catch(error => {
                            console.error('Unable to get user info', error);

                        });
                },
                failureCallback: (error) => {
                    microsoftTeams.appInitialization.notifyFailure({
                        reason: microsoftTeams.appInitialization.FailedReason.AuthFailed,
                        error,
                    });
                }
            });
```

To see back-end API check **[GraphController.cs](https://github.com/ardacetinkaya/Demo.MSTeams.App/blob/main/HelloWorld/Controllers/GraphController.cs)**. Basically some custom **[GraphService.cs](https://github.com/ardacetinkaya/Demo.MSTeams.App/blob/main/HelloWorld/Graph/GraphService.cs)** is injected into controller with a authentication provider, **[GraphAuthenticator.cs](https://github.com/ardacetinkaya/Demo.MSTeams.App/blob/main/HelloWorld/Graph/GraphAuthenticator.cs)** with **[Microsoft.Graph APIs](https://www.nuget.org/packages/Microsoft.Graph)**

To authenticate for Graph API call, there are some different approaches and providers. You can check **[here](https://docs.microsoft.com/en-us/graph/sdks/choose-authentication-providers?tabs=CS)** for more detailed info. In this repository you can find additional provider approaches with a simple code but for this demo following **DelegateAuthenticationProvider()** is used with MS Teams' user token with **[Microsoft.Identity.Client](https://www.nuget.org/packages/Microsoft.Identity.Client/)** 

```csharp
     client = new GraphServiceClient(new DelegateAuthenticationProvider(
     async (requestMessage) =>
     {
         try
         {
             // "token" is MS Teams' user token
             var userAssertion = new UserAssertion(token, "urn:ietf:params:oauth:grant-type:jwt-bearer");

             var clientApplication = ConfidentialClientApplicationBuilder.Create(_clientId)
                  .WithRedirectUri(_redirectUri)
                  .WithTenantId(_tenantId)
                  .WithClientSecret(_clientSecret)
                  .Build();

             var result = await clientApplication.AcquireTokenOnBehalfOf(_defaultScope, userAssertion)
                 .ExecuteAsync();

             requestMessage.Headers.Authorization =
                 new AuthenticationHeaderValue("Bearer", result.AccessToken);
         }
         catch (Exception ex)
         {
             Logger.LogError(ex, ex.Message);
             throw;

         }


     }));
```


To add an application to Microsoft Teams, a manifest file is required. This manifest file defines some properties for the application. You can check full manifest file description [here](https://docs.microsoft.com/en-us/microsoftteams/platform/resources/schema/manifest-schema).

For this repository check more basic and simple version ---> **[manifest.json](https://github.com/ardacetinkaya/Demo.MSTeams.App/blob/main/Manifest/manifest.json)**

# References

- [Calling Microsoft Graph from your Teams Application – Part 3: Tabs](https://bob1german.com/2020/08/31/calling-microsoft-graph-from-your-teams-application-part3/) 
  - _Also check Bob German's other posts, they are really good_
- [Authentication in Teams tabs using Microsoft Graph Toolkit](https://quickbites.dev/2020/07/29/auth-mgt-teams-tab/)
- [Choose a Microsoft Graph authentication provider based on scenario](https://docs.microsoft.com/en-us/graph/sdks/choose-authentication-providers?tabs=CS)
- [Create your first Teams app using C#](https://docs.microsoft.com/en-us/microsoftteams/platform/tutorials/get-started-dotnet-app-studio)
