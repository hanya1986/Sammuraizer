# Sammuraizer

## Prerequisites

To use this application, you need the following:
* [Node.js](https://nodejs.org/). Node is required to run the sample on a development server and to install dependencies. 

* [Bower](https://bower.io). Bower is required to install front-end dependencies.

* Either a [Microsoft account](https://www.outlook.com) or [Office 365 for business account](https://msdn.microsoft.com/en-us/office/office365/howto/setup-development-environment#bk_Office365Account)

## Register the application

1. Sign into the [App Registration Portal](https://apps.dev.microsoft.com/) using either your personal or work or school account.

2. Choose **Add an app**.

3. Enter a name for the app, and choose **Create application**. 
	
   The registration page displays, listing the properties of your app.

4. Copy the Application Id. This is the unique identifier for your app. 

5. Under **Platforms**, choose **Add Platform**.

6. Choose **Web**.

7. Make sure the **Allow Implicit Flow** check box is selected, and enter *http://localhost:8080/login* as the Redirect URI. 

8. Choose **Save**.


## Build and run the sample

1. Download or clone the Microsoft Graph Connect Sample for AngularJS.

2. Using your favorite IDE, open **config.js** in *public/scripts*.

3. Replace the **clientId** placeholder value with the application ID of your registered Azure application.

4. In a command prompt, run the following command in the root directory. This installs project dependencies, including the [HelloJS](http://adodson.com/hello.js/) client-side authententication library and the Microsoft Graph JavaScript SDK.

  ```
npm install
bower install
  ```
  
5. Run `npm start` to start the development server.

6. Navigate to `http://localhost:8080` in your web browser.

7. Choose the **Connect** button.

8. Sign in with your personal or work or school account and grant the requested permissions.

9. The HighLight Key Word feature currently is not working.

## Contributing

If you'd like to contribute to this sample, see [CONTRIBUTING.MD](/CONTRIBUTING.md).

This project has adopted the [Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/). For more information see the [Code of Conduct FAQ](https://opensource.microsoft.com/codeofconduct/faq/) or contact [opencode@microsoft.com](mailto:opencode@microsoft.com) with any additional questions or comments.
  
## Additional resources

- [Other Microsoft Graph Connect samples](https://github.com/MicrosoftGraph?utf8=%E2%9C%93&query=-Connect)
- [Microsoft Graph](http://graph.microsoft.io)

## Copyright
Copyright (c) 2016 Microsoft. All rights reserved.





