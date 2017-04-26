# Call Office 365 Services in Visual Studio with the Microsoft Graph

This walkthrough will show you how to use the Connected Services in Visual Studio to call the Microsoft Graph by getting a signed in user's profile photo, uploading their photo to their onedrive, and sending an email with a link to that photo.

## Getting Setup

1. To use the Office 365 Connected Services with the Microsoft Graph, you will need to download the [Visual Studio 2017 Preview](https://www.visualstudio.com/vs/preview/).

    > Even if you are using an older version of Visual Studio, Visual Studio 2017 Preview will work fine side by side.

2. You will also need an Office 365 Subscription. You can get one through your MSDN Subscription or by joining the [Office 365 Developer program](https://dev.office.com/devprogram)


## Get the Starter Project

Clone the following repository.  This sample will give you the references you need to authenticate against the Microsoft Graph.


## Add the Connected Service

Once you've setup your development environment, you're ready to add the Microsoft Graph service to your Visual Studio project. In order to successfully build this example, you'll need to have set the following permissions for the Office 365 Services with the Microsoft Graph. 
- For the *File* APIs, set permissions to **Have full access to your files**.
- For the *Mail* APIs, set permissions to **Send mail as you**.
- For the *User* APIs, set permissions to **Sign you in and read your profile**.

To do this, in Solution Explorer, right-click the project and select Add > Connected Service. Choose the Office 365 Connected Services and follow the wizard and select the scopes above.  You can also follow this same step to change the permisssions at a later time.

To verify, After you complete this, you should see the following set of resources in Solution Explorer
[ADD SCREENSHOT]

## Calling the Microsoft Graph

The Starter sample is configured to send a simple email. Let's now leverage power of the Microsoft Graph and update it to call multiple scenarios to complete our scenario.

1. Navigate to the 'GraphService.cs' which hosts our code to call the Graph.

2. **Uncomment** the following methods, which will get the User's Profile Photo (or use a cute picture of puppies if none is available), upload to OneDrive, and create a sharing link, respectively.  These methods use the Microsoft Graph SDK to get and set data in Office 365.

    ``` C#
        GetCurrentUserPhotoStreamAsync(GraphServiceClient graphClient)
    ```
    
    ``` C#
        UploadFileToOneDrive(GraphServiceClient graphClient, byte[] file)
    ```

    ```C#
        GetSharingLinkAsync(GraphServiceClient graphClient, string Id)
    ```
  
3. Now let's update the code that sends our email message out. 

- In Replace this line


- with the following


## Run the Sample
Press F5.  Click the Sign-in link on the top right -> then click the Send Email button.

You should see an email sent which will have a link to your profile photo.

-------------------------------------

## Drill in

You should now have been able to use Visual Studio 2017 to connect and configure your services.  The starter sample had scaffolding and references created for you.  

### Code of note

- [Startup.Auth.cs](/Microsoft%20Graph%20SDK%20ASPNET%20Connect/Microsoft%20Graph%20SDK%20ASPNET%20Connect/App_Start/Startup.Auth.cs). Authenticates the current user and initializes the sample's token cache.

- Models\[SessionTokenCache.cs](/Microsoft%20Graph%20SDK%20ASPNET%20Connect/Microsoft%20Graph%20SDK%20ASPNET%20Connect/TokenStorage/SessionTokenCache.cs). Stores the user's token information. You can replace this with your own custom token cache. Learn more in [Caching access tokens in a multitenant application](https://azure.microsoft.com/en-us/documentation/articles/guidance-multitenant-identity-token-cache/).

- Models\[SampleAuthProvider.cs](/Microsoft%20Graph%20SDK%20ASPNET%20Connect/Microsoft%20Graph%20SDK%20ASPNET%20Connect/Helpers/SampleAuthProvider.cs). Implements the local IAuthProvider interface, and gets an access token. 

- [SDKHelper.cs](/Microsoft%20Graph%20SDK%20ASPNET%20Connect/Microsoft%20Graph%20SDK%20ASPNET%20Connect/Helpers/SDKHelper.cs). Initializes the **GraphServiceClient** from the [Microsoft Graph .NET Client Library](https://github.com/microsoftgraph/msgraph-sdk-dotnet) that's used to interact with the Microsoft Graph.

- [HomeController.cs](/Microsoft%20Graph%20SDK%20ASPNET%20Connect/Microsoft%20Graph%20SDK%20ASPNET%20Connect/Controllers/HomeController.cs). Contains methods that use the **GraphServiceClient** to build and send calls to the Microsoft Graph service and to process the response.

- [Graph.cshtml](/Microsoft%20Graph%20SDK%20ASPNET%20Connect/Microsoft%20Graph%20SDK%20ASPNET%20Connect/Views/Home/Graph.cshtml). Contains the sample's UI. 


## Stuck?

Ping us on [StackOverflow](https://stackoverflow.com/questions/tagged/microsoftgraph?sort=newest)

