using SharePointASPCoreFileProvider.Models;
using Microsoft.AspNetCore.Cors;
using Microsoft.AspNetCore.Mvc;
using System.Collections.Generic;
using Syncfusion.Web.FileManager.Base;
using Microsoft.AspNetCore.Http.Features;
using Microsoft.AspNetCore.Http;
using System;
using Microsoft.Extensions.Configuration;
using System.Text.Json;

namespace SharePointASPCoreFileProvider.Controllers
{
	[Route("api/[controller]")]
	[EnableCors("AllowAllOrigins")]
	public class SharePointController : Controller
	{
		public SharePointProvider operation;

		public SharePointController(IConfiguration configuration)
		{
			this.operation = new SharePointProvider();
			var sharePointSettings = configuration.GetSection("SharePointSettings");
			this.operation.userSiteName = sharePointSettings["UserSiteName"];
			this.operation.userDriveId = sharePointSettings["UserDriveId"];
			this.operation.RegisterSharePoint(sharePointSettings["TenantId"], sharePointSettings["ClientId"], sharePointSettings["ClientSecret"]);
		}

		[Route("SharePointFileOperations")]
		public object FileOperations([FromBody] FileManagerDirectoryContent args)
		{
			if (args.Action == "delete" || args.Action == "rename")
			{
				if ((args.TargetPath == null) && (args.Path == ""))
				{
					FileManagerResponse response = new FileManagerResponse();
					response.Error = new ErrorDetails { Code = "401", Message = "Restricted to modify the root folder." };
					return this.operation.ToCamelCase(response);
				}
			}
			switch (args.Action)
			{
				case "read":
					// reads the file(s) or folder(s) from the given path.
					return this.operation.ToCamelCase(this.operation.GetFiles(args.Path, args.ShowHiddenItems));
				case "delete":
					// deletes the selected file(s) or folder(s) from the given path.
					return this.operation.ToCamelCase(this.operation.Delete(args.Path, args.Names, args.Data));
				case "copy":
					// copies the selected file(s) or folder(s) from a path and then pastes them into a given target path.
					return this.operation.ToCamelCase(this.operation.Copy(args.Action, args.Path, args.TargetPath, args.Names, args.RenameFiles, args.TargetData, args.Data));
				case "move":
					// cuts the selected file(s) or folder(s) from a path and then pastes them into a given target path.
					return this.operation.ToCamelCase(this.operation.Move(args.Action, args.Path, args.TargetPath, args.Names, args.RenameFiles, args.TargetData, args.Data));
				case "details":
					// gets the details of the selected file(s) or folder(s).
					return this.operation.ToCamelCase(this.operation.Details(args.Path, args.Names, args.Data));
				case "create":
					// creates a new folder in a given path.
					return this.operation.ToCamelCase(this.operation.Create(args.Path, args.Name, args.Data));
				case "search":
					// gets the list of file(s) or folder(s) from a given path based on the searched key string.
					return this.operation.ToCamelCase(this.operation.Search(args.Path, args.SearchString, args.ShowHiddenItems, args.CaseSensitive, args.Data));
				case "rename":
					// renames a file or folder.
					return this.operation.ToCamelCase(this.operation.Rename(args.Path, args.Name, args.NewName, false, args.ShowFileExtension, args.Data));
			}
			return null;
		}

		[Route("SharePointUpload")]
		public IActionResult SharePointUpload(string path, IList<IFormFile> uploadFiles, string action, string data)
		{
			FileManagerResponse uploadResponse;
			FileManagerDirectoryContent[] dataObject = new FileManagerDirectoryContent[1];
			var options = new JsonSerializerOptions { PropertyNamingPolicy = null, };
			dataObject[0] = JsonSerializer.Deserialize<FileManagerDirectoryContent>(data, options);
			uploadResponse = this.operation.Upload(path, uploadFiles, action, dataObject);
			if (uploadResponse.Error != null)
			{
				Response.Clear();
				Response.ContentType = "application/json; charset=utf-8";
				Response.StatusCode = Convert.ToInt32(uploadResponse.Error.Code);
				Response.HttpContext.Features.Get<IHttpResponseFeature>().ReasonPhrase = uploadResponse.Error.Message;
			}
			return Content("");
		}

		[Route("SharePointDownload")]
		public IActionResult Download(string downloadInput)
		{
			var options = new JsonSerializerOptions
			{
				PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
			};
			FileManagerDirectoryContent args = JsonSerializer.Deserialize<FileManagerDirectoryContent>(downloadInput, options);
			return this.operation.Download(args.Path, args.Names, args.Data);
		}

		[Route("SharePointGetImage")]
		public IActionResult GetImage(FileManagerDirectoryContent args)
		{
			return this.operation.GetImage(args.Path, args.Id, false, args.Data);
		}
	}
}
