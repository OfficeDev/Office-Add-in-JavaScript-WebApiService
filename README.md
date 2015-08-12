# Office add-in: Create a web service for an Office add-in using the ASP.NET Web API

**Table of contents**

* [Summary](#summary)
* [Prerequisites](#prerequisites)
* [Key components of the sample](#components)
* [Description of the code](#codedescription)
* [Build and debug](#build)
* [Troubleshooting](#troubleshooting)
* [Questions and comments](#questions)
* [Additional resources](#additional-resources)

<a name="summary"></a>
##Summary
This sample demonstrates how to create and query an ASP.NET Web API service from an apps for Office. The example app is comprised of a "Send Feedback" page, which lets a user submit feedback, and uses a Web API service to send it to the developer team. 

<a name="prerequisites"></a>
## Prerequisites ##

This sample requires the following:  

  - Excel 2013, Word 2013, PowerPoint 2013, or Project 2013
  - Visual Studio 2013 with Update 5 or Visual Studio 2015.  
  - Any browser that supports ECMAScript 5.1, HTML5, and CSS3, such as Internet Explorer 9, Chrome 13, Firefox 5, Safari 5.0.6, or a later version of these browsers.
  - Familiarity with JavaScript programming and web services.

<a name="components"></a>
## Key components of the sample
The sample solution contains the following files:

- WebApi SampleManifest.xml: The manifest file for the Office add-in.
- App\Home\Home.html: The HTML user interface for the mail add-in for Office.
- App\Home\Home.js: The JavaScript file that handles sending the attachment information to the remote Attachment service included with this sample.
- Controllers\SendFeedbackController.cs: The business logic for the sample Web API service.
- App_Start\WebApiConfig.cs: Binds the default routing for the Web API service.

<a name="codedescription"></a>
##Description of the code
The Office add-in makes an AJAX request to the web service, passing in data from the client-side JavaScript code. The Web API controller receives the data, performs an action, and returns the results back to the caller. The AJAX call then completes, 
 displaying the results or showing an error message.


<a name="build"></a>
## Build and debug ##
**Note**: The mail add-in is activated on any email message in the user's Inbox. You can make it easier to test the add-in by sending one or more email messages to your test account before you run the sample add-in.

The sample will run right out of the box, but it won't be able to send feedback unless you configure appropriate credentials in the "SendFeedbackController.cs" file (in the Controllers folder of the web project). Update the following constants with actual values:

    ```c#
	const string MailingAddressFrom = "app_name@contoso.com ";
    const string MailingAddressTo = "dev_team@contoso.com";
    const string SmtpHost = "smtp.contoso.com";
    const int SmtpPort = 587;
    const bool SmtpEnableSsl = true;
    const string SmtpCredentialsUsername = "username";
    const string SmtpCredentialsPassword = "password";
	```

1. Open the solution in Visual Studio.
2. Press F5 to build and deploy the sample add-in to the client that's specified as the start document (by default, Excel). To change this setting in the Properties view (View > Properties Window), click the WebApi Sample project in Solution Explorer and select your preferred client.
3. Choose a rating in the drop-down list, enter some feedback, and click **Send it!** A toast notification opens to tell you whether your feedback was successfully sent.


<a name="troubleshooting"></a>
##Troubleshooting

- If the add-in fails to send feedback (shows a notification message with "Sorry, your feedback could not be sent"), check that you configured an appropriate email address in SendFeedbackController.cs. Alternatively, you can remove the mail-sending code, and/or replace it with a different form of sending feedback (e.g., logging to a database).

<a name="questions"></a>
##Questions and comments##

- If you have any trouble running this sample, please [log an issue](https://github.com/OfficeDev/Office-Add-in-JavaScript-WebApiService/issues).
- Questions about Office Add-in development in general should be posted to [Stack Overflow](http://stackoverflow.com/questions/tagged/office-addins). Make sure that your questions or comments are tagged with [office-addins].

<a name="additional-resources"></a>
## Additional resources ##

- [Web API: The Official Microsoft ASP.NET Site](http://www.asp.net/web-api)
- [Walkthrough: Create a web service for an app for Office using the ASP.NET Web API](http://blogs.msdn.com/b/officeapps/archive/2013/06/05/create-a-web-service-for-an-app-for-office-using-the-asp-net-web-api.aspx) (applies to an earlier version of this sample)

## Copyright
Copyright (c) 2015 Microsoft. All rights reserved.
