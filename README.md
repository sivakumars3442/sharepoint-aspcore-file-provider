# sharepoint-aspcore-file-provider

This repository contains the ASP.NET Core SharePoint file system providers for the Syncfusion File Manager component.

## Key Features

The SharePoint file system provider offers file system support for the FileManager component with Microsoft SharePoint.

The following actions can be performed with the SharePoint file system provider:

| **Actions** | **Description** |
| --- | --- |
| Read     | Reads the files from the SharePoint document library. |
| Details  | Provides details about files Type, Size, Location and Modified date. |
| Download | Downloads the selected file or folder from the SharePoint document library. |
| Upload   | Uploads files to the SharePoint document library. It accepts uploaded media with the following characteristics: <ul><li>Maximum file size: 30MB</li><li>Accepted Media MIME types: */*</li></ul> |
| Create   | Creates a new folder in the SharePoint document library. |
| Delete   | Removes a file or folder from the SharePoint document library. |
| Copy     | Copies the selected files to the target location within the SharePoint document library. |
| Move     | Moves the copied files to the desired location within the SharePoint document library. |
| Rename   | Renames a folder or file in the SharePoint document library. |
| Search   | Searches for a file or folder in the SharePoint document library. |

## Prerequisites

To set up the SharePoint service provider, follow these steps:

1. **Create an App Registration in Azure Active Directory (AAD):** 
   - Navigate to the Azure portal and create a new app registration under Azure Active Directory.
   - Note down the **Tenant ID**, **Client ID**, and **Client Secret** from the app registration.

2. **Use Microsoft Graph Instance:** 
   - With the obtained Tenant ID, Client ID, and Client Secret, you can create a Microsoft Graph instance.
   - This instance will be used to interact with the SharePoint document library.

3. **Use Details from `appsettings.json`:**
   - The `SharePointController` is already configured to use the credentials provided in the `appsettings.json` file.
   - You only need to provide your `Tenant ID`, `Client ID`, `Client Secret`, `User Site Name`, and `User Drive ID` in the `appsettings.json` file, and the application will automatically initialize the SharePoint service.

   ### Example `appsettings.json` Configuration

```json
{
  "Logging": {
    "LogLevel": {
      "Default": "Warning"
    }
  },
  "SharePointSettings": {
    "TenantId": "<--Tenant Id-->",
    "ClientId": "<--Client Id-->",
    "ClientSecret": "<--Client Secret-->",
    "UserSiteName": "<--User Site Name-->",
    "UserDriveId": "<--User Drive ID-->"
  },
  "AllowedHosts": "*"
}
```

Replace "<--User Site Name-->", "<--User Drive ID-->", "tenantId", "clientId", and "clientSecret" with your actual values.

## How to run this application?

* Checkout this project to a location in your disk.
* Open the solution file using Visual Studio 2022.
* Restore the NuGet packages by rebuilding the solution.
* Run the project.

## Running application

Once cloned, open solution file in visual studio.Then build the project after restoring the nuget packages and run it.


## File Manager AjaxSettings

To access the basic actions such as Read, Delete, Copy, Move, Rename, Search, and Get Details of File Manager using the SharePoint service, map the following code snippet in the `AjaxSettings` property of File Manager. Please refer to the platform-specific documentation for more details:

Here, the `hostUrl` will be your locally hosted port number.

```
  var hostUrl = http://localhost:62870/;
  ajaxSettings: {
        url: hostUrl + 'api/SharePointProvider/SharePointFileOperations'
  }
```

Refer to the below File Manager UG Documentation:

- JavaScript – https://ej2.syncfusion.com/javascript/documentation/file-manager/es5-getting-started#using-cdn-link-for-script-and-style-reference
- TypeScrip – https://ej2.syncfusion.com/documentation/file-manager/getting-started#initialize-the-file-manager
- React – https://ej2.syncfusion.com/react/documentation/file-manager/getting-started#run-the-application
- Vue – https://ej2.syncfusion.com/vue/documentation/file-manager/getting-started#add-syncfusion-vue-component
- Angular – https://ej2.syncfusion.com/angular/documentation/file-manager/getting-started#adding-syncfusion-component
- Blazor – https://blazor.syncfusion.com/documentation/file-manager/getting-started-with-server-app#add-blazor-filemanager-component
- ASP.NET Core – https://ej2.syncfusion.com/aspnetcore/documentation/file-manager/getting-started#add-aspnet-core-filemanager-control
- ASP.NET MVC – https://ej2.syncfusion.com/aspnetmvc/documentation/file-manager/getting-started#add-aspnet-mvc-filemanager-control

## File download AjaxSettings

To perform download operation, initialize the `downloadUrl` property in ajaxSettings of the File Manager component.

```
  var hostUrl = http://localhost:62870/;
  ajaxSettings: {
    url: hostUrl + 'api/SharePointProvider/SharePointFileOperations',
    downloadUrl: hostUrl + 'api/SharePointProvider/SharePointDownload'
  }
```

- JavaScript – https://ej2.syncfusion.com/javascript/documentation/file-manager/es5-getting-started#file-download-support
- TypeScrip – https://ej2.syncfusion.com/documentation/file-manager/getting-started#file-download-support
- React – https://ej2.syncfusion.com/react/documentation/file-manager/getting-started#file-download-support
- Vue – https://ej2.syncfusion.com/vue/documentation/file-manager/getting-started#file-download-support
- Angular – https://ej2.syncfusion.com/angular/documentation/file-manager/getting-started#file-download-support
- Blazor – https://blazor.syncfusion.com/documentation/file-manager/getting-started-with-server-app#file-download-support
- ASP.NET Core – https://ej2.syncfusion.com/aspnetcore/documentation/file-manager/getting-started#file-download-support
- ASP.NET MVC – https://ej2.syncfusion.com/aspnetmvc/documentation/file-manager/getting-started#file-download-support

## File upload AjaxSettings

To perform upload operation, initialize the `uploadUrl` property in ajaxSettings of the File Manager component.

```
  var hostUrl = http://localhost:62870/;
  ajaxSettings: {
    url: hostUrl + 'api/SharePointProvider/SharePointFileOperations',
    uploadUrl: hostUrl + 'api/SharePointProvider/SharePointUpload'
  }
```

- JavaScript – https://ej2.syncfusion.com/javascript/documentation/file-manager/es5-getting-started#file-upload-support
- TypeScrip – https://ej2.syncfusion.com/documentation/file-manager/getting-started#file-upload-support
- React – https://ej2.syncfusion.com/react/documentation/file-manager/getting-started#file-upload-support
- Vue – https://ej2.syncfusion.com/vue/documentation/file-manager/getting-started#file-upload-support
- Angular – https://ej2.syncfusion.com/angular/documentation/file-manager/getting-started#file-upload-support
- Blazor – https://blazor.syncfusion.com/documentation/file-manager/getting-started-with-server-app#file-upload-support
- ASP.NET Core – https://ej2.syncfusion.com/aspnetcore/documentation/file-manager/getting-started#file-upload-support
- ASP.NET MVC – https://ej2.syncfusion.com/aspnetmvc/documentation/file-manager/getting-started#file-upload-support

## File image preview AjaxSettings

To perform image preview support in the File Manager component, initialize the `getImageUrl` property in ajaxSettings of the File Manager component.

```
  var hostUrl = http://localhost:62870/;
  ajaxSettings: {
    url: hostUrl + 'api/SharePointProvider/SharePointFileOperations',
    getImageUrl: hostUrl + 'api/SharePointProvider/SharePointGetImage'
  }
```

- JavaScript – https://ej2.syncfusion.com/javascript/documentation/file-manager/es5-getting-started#image-preview-support
- TypeScrip – https://ej2.syncfusion.com/documentation/file-manager/getting-started#image-preview-support
- React – https://ej2.syncfusion.com/react/documentation/file-manager/getting-started#image-preview-support
- Vue – https://ej2.syncfusion.com/vue/documentation/file-manager/getting-started#image-preview-support
- Angular – https://ej2.syncfusion.com/angular/documentation/file-manager/getting-started#image-preview-support
- Blazor – https://blazor.syncfusion.com/documentation/file-manager/getting-started-with-server-app#image-preview-support
- ASP.NET Core – https://ej2.syncfusion.com/aspnetcore/documentation/file-manager/getting-started#image-preview-support
- ASP.NET MVC – https://ej2.syncfusion.com/aspnetmvc/documentation/file-manager/getting-started#image-preview-support

## Support

Product support is available for through following mediums.

* Creating incident in Syncfusion [Direct-trac](https://www.syncfusion.com/support/directtrac/incidents?utm_source=npm&utm_campaign=filemanager) support system or [Community forum](https://www.syncfusion.com/forums/essential-js2?utm_source=npm&utm_campaign=filemanager).
* New [GitHub issue](https://github.com/syncfusion/ej2-javascript-ui-controls/issues/new).
* Ask your query in [Stack Overflow](https://stackoverflow.com/?utm_source=npm&utm_campaign=filemanager) with tag `syncfusion`, `syncfusion-blazor` and `ej2`.

## License

Check the license detail [here](https://github.com/syncfusion/ej2-javascript-ui-controls/blob/master/license).