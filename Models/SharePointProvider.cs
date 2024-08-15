using Azure.Core;
using Microsoft.Graph;
using Azure.Identity;
using Microsoft.Graph.Models;
using Syncfusion.Web.FileManager.Base;
using System.IO;
using Microsoft.Graph.Drives.Item.Items.Item.Copy;
using Microsoft.AspNetCore.Mvc;
using System.IO.Compression;
using System;
using Microsoft.AspNetCore.Http;
using System.Collections.Generic;
using System.Threading.Tasks;
using System.Linq;
using Newtonsoft.Json.Serialization;
using Newtonsoft.Json;

namespace SharePointASPCoreFileProvider.Models
{
	public class SharePointProvider
	{
		public TokenCredential _tokenCredential { get; set; } = default;
		public string[] _scopes { get; set; } = new[] { "https://graph.microsoft.com/.default" };
		public string userSiteName { get; set; } = string.Empty;
		public string userDriveId { get; set; } = string.Empty;
		public GraphServiceClient graphServiceClient { get; set; } = default;
		public AccessDetails AccessDetails { get; set; } = new AccessDetails();
		private string accessMessage { get; set; } = string.Empty;

		public void RegisterSharePoint(string tenantId, string clientId, string clientSecret)
		{
			_tokenCredential = new ClientSecretCredential(tenantId, clientId, clientSecret);
			graphServiceClient = new GraphServiceClient(_tokenCredential, _scopes);
		}

		public async Task GetSharePointDrive()
		{
			SiteCollectionResponse userSite = new SiteCollectionResponse();
			userSite = await graphServiceClient.Sites?.GetAsync(requestConfiguration =>
								requestConfiguration.QueryParameters.Search = userSiteName);
			Drive userDrive = new Drive();
			userDrive = await graphServiceClient.Sites[userSite?.Value?.FirstOrDefault()?.Id]?.Drive.GetAsync();
			userDriveId = userDrive?.Id;
		}

		public virtual FileManagerResponse GetFiles(string path, bool showHiddenItems, params FileManagerDirectoryContent[] data)
		{
			return GetFilesAsync(path, showHiddenItems, data).Result;
		}

		public async Task<FileManagerResponse> GetFilesAsync(string path, bool showHiddenItems, params FileManagerDirectoryContent[] data)
		{
			if (string.IsNullOrEmpty(userDriveId)) 
			{
				await GetSharePointDrive();
			}
			FileManagerDirectoryContent cwd = new FileManagerDirectoryContent();
			List<FileManagerDirectoryContent> child = new List<FileManagerDirectoryContent>();
			FileManagerResponse readResponse = new FileManagerResponse();
			try
			{
				DriveItem root = new DriveItem();
				if (path == "/")
				{
					root = await graphServiceClient.Drives[userDriveId].Items["root"].GetAsync();
				}
				else
				{
					var filterId = path.Split(new char[] { '/' }, StringSplitOptions.RemoveEmptyEntries);
					root = await graphServiceClient.Drives[userDriveId].Items[filterId.LastOrDefault()].GetAsync();
				}
				cwd.Name = root?.Name;
				cwd.Id = root?.Id;
				cwd.ParentId = root?.ParentReference != null ? root.ParentReference.Id : null;
				cwd.Size = (long)(root?.Size);
				cwd.IsFile = root.Folder == null ? true : false;
				cwd.DateModified = root.LastModifiedDateTime.Value.DateTime;
				cwd.DateCreated = root.CreatedDateTime.Value.DateTime;
				cwd.HasChild = root.File != null ? false : await HasChild(root);
				cwd.Type = Path.GetExtension(root.Name);
				cwd.FilterId = await GetFilterId(root);
				cwd.FilterPath = await GetFilterPath(root);
				cwd.Permission = GetPathPermission(cwd.FilterPath + (cwd.IsFile ? cwd.Name : Path.GetFileNameWithoutExtension(cwd.Name)), cwd.IsFile);
				child = await GetChild(root.Id);
				readResponse.CWD = cwd;
				if ((cwd.Permission != null && !cwd.Permission.Read))
				{
					readResponse.Files = null;
					accessMessage = cwd.Permission.Message;
					throw new UnauthorizedAccessException("'" + cwd.Name + "' is not accessible. You need permission to perform the read action.");
				}
				readResponse.Files = child;
				return readResponse;
			}
			catch (Exception e)
			{
				ErrorDetails er = new ErrorDetails();
				er.Message = e.Message.ToString();
				er.Code = er.Message.Contains("is not accessible. You need permission") ? "401" : "417";
				if ((er.Code == "401") && !string.IsNullOrEmpty(accessMessage)) { er.Message = accessMessage; }
				readResponse.Error = er;
				return readResponse;
			}
		}

		public FileManagerResponse Delete(string path, string[] names, params FileManagerDirectoryContent[] data)
		{
			return AsyncDelete(path, names, data).Result;
		}

		public async Task<FileManagerResponse> AsyncDelete(string path, string[] names, params FileManagerDirectoryContent[] data)
		{
			if (string.IsNullOrEmpty(userDriveId))
			{
				await GetSharePointDrive();
			}
			FileManagerResponse removeResponse = new FileManagerResponse();
			try
			{
				List<FileManagerDirectoryContent> deletedItems = new List<FileManagerDirectoryContent>();
				if(graphServiceClient != null)
				{
					foreach(FileManagerDirectoryContent item in data)
					{
						AccessPermission PathPermission = GetPathPermission(item.FilterPath + item.Name, item.IsFile);
						if (PathPermission != null && (!PathPermission.Read || !PathPermission.Write))
						{
							accessMessage = PathPermission.Message;
							throw new UnauthorizedAccessException("'" + item.Name + "' is not accessible.  You need permission to perform the write action.");
						}
						deletedItems.Add(item);
						await graphServiceClient.Drives[userDriveId].Items[item.Id].DeleteAsync();
					}
				}
				removeResponse.Files = deletedItems;
				return removeResponse;
			}
			catch (Exception ex)
			{
				ErrorDetails er = new ErrorDetails();
				er.Message = ex.Message.ToString();
				er.Code = er.Message.Contains(" is not accessible.  You need permission") ? "401" : "417";
				if ((er.Code == "401") && !string.IsNullOrEmpty(accessMessage)) { er.Message = accessMessage; }
				removeResponse.Error = er;
				return removeResponse;
			}
		}

		public FileManagerResponse Copy(string action, string path, string targetPath, string[] names, string[] replacedItemNames, FileManagerDirectoryContent TargetData, params FileManagerDirectoryContent[] data)
		{
			return this.CopyOrMoveDriveItemAsync(action, path, targetPath, names, replacedItemNames, TargetData, data).Result;
		}

		public FileManagerResponse Move(string action, string path, string targetPath, string[] names, string[] replacedItemNames, FileManagerDirectoryContent TargetData, params FileManagerDirectoryContent[] data)
		{
			return this.CopyOrMoveDriveItemAsync(action, path, targetPath, names, replacedItemNames, TargetData, data).Result;
		}

		public async Task<FileManagerResponse> CopyOrMoveDriveItemAsync(string action, string path, string targetPath, string[] names, string[] replacedItemNames, FileManagerDirectoryContent TargetData, params FileManagerDirectoryContent[] data)
		{
			if (string.IsNullOrEmpty(userDriveId))
			{
				await GetSharePointDrive();
			}
			FileManagerResponse moveResponse = new FileManagerResponse();
			List<FileManagerDirectoryContent> transferredList = new List<FileManagerDirectoryContent>();
			try
			{

				foreach (FileManagerDirectoryContent item in data)
				{
					var targetChildren = await this.graphServiceClient.Drives[userDriveId].Items[TargetData.Id].Children.GetAsync();
					AccessPermission PathPermission = GetPathPermission(item.FilterPath + item.Name, item.IsFile);
					if (action == "copy")
					{
						if (PathPermission != null && (!PathPermission.Read || !PathPermission.Copy))
						{
							accessMessage = PathPermission.Message;
							throw new UnauthorizedAccessException("'" + item.Name + "' is not accessible. You need permission to perform the copy action.");
						}
						var requestBody = new CopyPostRequestBody
						{
							ParentReference = new ItemReference
							{
								DriveId = userDriveId,
								Id = TargetData.Id
							},
							Name = item?.Name
						};
						await this.graphServiceClient.Drives[userDriveId].Items[item.Id].Copy.PostAsync(requestBody);
					}
					else
					{
						if (PathPermission != null && (!PathPermission.Read || !PathPermission.Write))
						{
							accessMessage = PathPermission.Message;
							throw new UnauthorizedAccessException("'" + item.Name + "' is not accessible. You need permission to perform the Write action.");
						}
						var moveRequestBody = new DriveItem
						{
							ParentReference = new ItemReference
							{
								Id = TargetData.Id,
							},
							Name = item.Name,
						};
						await this.graphServiceClient.Drives[userDriveId].Items[item.Id].PatchAsync(moveRequestBody);
					}
				}
				foreach(FileManagerDirectoryContent item in data)
				{
					DriveItem transferredDrive = null;
					while (transferredDrive == null)
					{
						try
						{
							transferredDrive = await this.graphServiceClient.Drives[userDriveId].Items[TargetData.Id].Children[item.Name].GetAsync();
						}
						catch (Exception)
						{
							continue;
						}
						if (transferredDrive != null)
						{
							FileManagerDirectoryContent file = new FileManagerDirectoryContent();
							file.Name = transferredDrive.Name;
							file.Id = transferredDrive.Id;
							file.ParentId = transferredDrive.ParentReference?.Id;
							file.Size = transferredDrive.Size ?? 0;
							file.IsFile = transferredDrive.Folder == null ? true : false;
							file.DateModified = transferredDrive.LastModifiedDateTime.Value.DateTime;
							file.DateCreated = transferredDrive.CreatedDateTime.Value.DateTime;
							file.HasChild = transferredDrive.Folder != null && transferredDrive.Folder.ChildCount > 0 ? true : false;
							file.Type = Path.GetExtension(transferredDrive.Name);
							file.FilterId = await GetFilterId(transferredDrive);
							file.FilterPath = await GetFilterPath(transferredDrive);
							file.Permission = GetPathPermission(file.FilterPath + (file.IsFile ? file.Name : Path.GetFileNameWithoutExtension(file.Name)), file.IsFile);
							transferredList.Add(file);
						}
					}
				}
				moveResponse.Files = transferredList;
				return moveResponse;
			}
			catch (Exception ex)
			{
				ErrorDetails error = new ErrorDetails();
				error.Message = ex.Message.ToString();
				error.Code = error.Message.Contains("is not accessible. You need permission") ? "401" : "404";
				if ((error.Code == "401") && !string.IsNullOrEmpty(accessMessage)) { error.Message = accessMessage; }
				error.FileExists = moveResponse.Error?.FileExists;
				moveResponse.Error = error;
				return moveResponse;
			}
		}

		public FileManagerResponse Details(string path, string[] names, params FileManagerDirectoryContent[] data)
		{
			return this.DetailsAsync(path, names, data).Result;
		}

		public async Task<FileManagerResponse> DetailsAsync(string path, string[] names, params FileManagerDirectoryContent[] data)
		{
			if (string.IsNullOrEmpty(userDriveId))
			{
				await GetSharePointDrive();
			}
			FileManagerResponse detailsResponse = new FileManagerResponse();
			Syncfusion.Web.FileManager.Base.FileDetails detailFiles = new Syncfusion.Web.FileManager.Base.FileDetails();
			try
			{
				if(data.Length == 0 || data.Length == 1)
				{
					var root = data.FirstOrDefault();
					detailFiles.Name = root.Name;
					detailFiles.IsFile = root.IsFile;
					detailFiles.Size = byteConversion(root.Size);
					detailFiles.Created = root.DateCreated;
					detailFiles.Location = (root.ParentId == null || root.ParentId == "/") ? "root" : "root"+root.FilterPath.TrimEnd('/') ;
					detailFiles.Modified = root.DateModified;
					detailFiles.Permission = root.Permission;
					detailFiles.MultipleFiles = false;
				}
				else
				{
					detailFiles.MultipleFiles = true;
					detailFiles.Name = string.Join(", ", data.Select(x => x.Name).ToArray());
					detailFiles.Location = (data.FirstOrDefault().ParentId == null || data.FirstOrDefault().ParentId == "/") ? "root" : "root" + data.FirstOrDefault().FilterPath.TrimEnd('/');
					long sizeValue = 0;
					foreach (var item in data)
					{
						sizeValue = sizeValue + item.Size;
					}
					detailFiles.Size = byteConversion(sizeValue);

				}
				detailsResponse.Details = detailFiles;
				return detailsResponse;
			}
			catch (Exception e)
			{
				ErrorDetails er = new ErrorDetails();
				er.Message = e.Message.ToString();
				er.Code = er.Message.Contains("is not accessible. You need permission") ? "401" : "417";
				detailsResponse.Error = er;
				return detailsResponse;
			}
		}

		public FileManagerResponse Create(string path, string name, params FileManagerDirectoryContent[] data)
		{
			return this.createAsync(path, name, data).Result;
		}

		public async Task<FileManagerResponse> createAsync(string path, string name, params FileManagerDirectoryContent[] data)
		{
			if (string.IsNullOrEmpty(userDriveId))
			{
				await GetSharePointDrive();
			}
			FileManagerResponse createResponse = new FileManagerResponse();
			try
			{
				AccessPermission PathPermission = GetPathPermission(data[0].FilterPath + data[0].Name, data[0].IsFile);
				var targetChildren = await this.graphServiceClient.Drives[userDriveId].Items[data.FirstOrDefault().Id].Children.GetAsync();
				if (targetChildren.Value.Any(targetItem => targetItem.Name.Equals(name, StringComparison.OrdinalIgnoreCase)))
				{
					ErrorDetails er = new ErrorDetails();
					er.Code = "400";
					er.Message = "A file or folder with the name " + name + " already exists.";
					createResponse.Error = er;
					return createResponse;
				}
				else
				{
					if (PathPermission != null && (!PathPermission.Read || !PathPermission.WriteContents))
					{
						accessMessage = PathPermission.Message;
						throw new UnauthorizedAccessException("'" + name + "' is not accessible. You need permission to perform the writeContents action.");
					}
					var folder = new DriveItem
					{
						Name = name,
						Folder = new Folder()
					};
					await this.graphServiceClient.Drives[userDriveId].Items[data.FirstOrDefault().Id].Children.PostAsync(folder);
					var createdItem = await this.graphServiceClient.Drives[userDriveId].Items[data.FirstOrDefault().Id].Children[name].GetAsync();
					FileManagerDirectoryContent file = new FileManagerDirectoryContent();
					file.Name = createdItem.Name;
					file.Id = createdItem.Id;
					file.ParentId = createdItem.ParentReference?.Id;
					file.Size = createdItem.Size ?? 0;
					file.IsFile = createdItem.Folder == null ? true : false;
					file.DateModified = createdItem.LastModifiedDateTime.Value.DateTime;
					file.DateCreated = createdItem.CreatedDateTime.Value.DateTime;
					file.HasChild = createdItem.Folder != null && createdItem.Folder.ChildCount > 0 ? true : false;
					file.Type = Path.GetExtension(createdItem.Name);
					file.FilterId = await GetFilterId(createdItem);
					file.FilterPath = await GetFilterPath(createdItem);
					file.Permission = GetPathPermission(file.FilterPath + (file.IsFile ? file.Name : Path.GetFileNameWithoutExtension(file.Name)), file.IsFile);
					createResponse.Files = new List<FileManagerDirectoryContent> { file };
				}
				return createResponse;
			}
			catch (Exception ex)
			{
				ErrorDetails er = new ErrorDetails();
				er.Message = ex.Message.ToString();
				er.Code = er.Message.Contains("is not accessible. You need permission") ? "401" : "417";
				if ((er.Code == "401") && !string.IsNullOrEmpty(accessMessage)) { er.Message = accessMessage; }
				createResponse.Error = er;
				return createResponse;
			}
		}

		public FileManagerResponse Search(string path, string searchString, bool showHiddenItems, bool caseSensitive, params FileManagerDirectoryContent[] data)
		{
			return this.SearchAsync(path, searchString, showHiddenItems, caseSensitive, data).Result;
		}

		public async Task<FileManagerResponse> SearchAsync(string path, string searchString, bool showHiddenItems, bool caseSensitive, params FileManagerDirectoryContent[] data)
		{
			if (string.IsNullOrEmpty(userDriveId))
			{
				await GetSharePointDrive();
			}
			FileManagerResponse searchResponse = new FileManagerResponse();
			try
			{
				AccessPermission permission = GetPathPermission(data[0].FilterPath + (data[0].IsFile ? data[0].Name : Path.GetFileNameWithoutExtension(data[0].Name)), data[0].IsFile);

				if ((permission != null && !permission.Read))
				{
					searchResponse.Files = null;
					accessMessage = data[0].Permission.Message;
					throw new UnauthorizedAccessException("'" + data[0].Name + "' is not accessible. You need permission to perform the read action.");
				}
				var searchResults = await this.graphServiceClient.Drives[userDriveId].SearchWithQ(searchString.Trim('*')).GetAsSearchWithQGetResponseAsync();
				searchResponse.Files = new List<FileManagerDirectoryContent>();
				searchResponse.Files = searchResults.Value.Where(item => item.Name.Equals(searchString.Trim('*'), caseSensitive ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase)).Select(item => new FileManagerDirectoryContent
				{
					Name = item.Name,
					Id = item.Id,
					ParentId = item.ParentReference?.Id,
					Size = item.Size ?? 0,
					IsFile = item.Folder == null ? true : false,
					DateModified = item.LastModifiedDateTime.Value.DateTime,
					DateCreated = item.CreatedDateTime.Value.DateTime,
					HasChild = item.Folder != null && item.Folder.ChildCount > 0 ? true : false,
					Type = Path.GetExtension(item.Name),
					FilterId = GetFilterId(item).Result,
					FilterPath = GetFilterPath(item).Result,
					Permission = GetPathPermission(item.ParentReference.Path + (item.Folder == null ? item.Name : Path.GetFileNameWithoutExtension(item.Name)), item.Folder == null ? true : false)
				}).ToList();
				searchResponse.CWD = data.FirstOrDefault();
				return searchResponse;
			}
			catch (Exception e)
			{
				ErrorDetails er = new ErrorDetails();
				er.Message = e.Message.ToString();
				er.Code = er.Message.Contains("is not accessible. You need permission") ? "401" : "417";
				if ((er.Code == "401") && !string.IsNullOrEmpty(accessMessage)) { er.Message = accessMessage; }
				searchResponse.Error = er;
				return searchResponse;
			}
		}

		public FileManagerResponse Rename(string path, string name, string newName, bool replace = false, bool showFileExtension = true, params FileManagerDirectoryContent[] data)
		{
			return this.RenameAsync(path, name, newName, replace, showFileExtension, data).Result;
		}

		public async Task<FileManagerResponse> RenameAsync(string path, string name, string newName, bool replace = false, bool showFileExtension = true, params FileManagerDirectoryContent[] data)
		{
			if (string.IsNullOrEmpty(userDriveId))
			{
				await GetSharePointDrive();
			}
			FileManagerResponse renameResponse = new FileManagerResponse();
			FileManagerDirectoryContent file = new FileManagerDirectoryContent();
			try
			{
				AccessPermission PathPermission = GetPathPermission(data[0].FilterPath + data[0].Name, data[0].IsFile);
				if (PathPermission != null && (!PathPermission.Read || !PathPermission.Write))
				{
					accessMessage = PathPermission.Message;
					throw new UnauthorizedAccessException();
				}
				var targetChildren = await this.graphServiceClient.Drives[userDriveId].Items[data.FirstOrDefault().ParentId].Children.GetAsync();
				if (targetChildren.Value.Any(targetItem => targetItem.Name.Equals(newName, StringComparison.OrdinalIgnoreCase)))
				{
					ErrorDetails er = new ErrorDetails();
					er.Code = "400";
					er.Message = "Cannot rename " + name + " to " + newName + ": destination already exists.";
					renameResponse.Error = er;
					return renameResponse;
				}
				else
				{
					var renameItem = new DriveItem
					{
						Name = newName
					};
					await this.graphServiceClient.Drives[userDriveId].Items[data.FirstOrDefault().Id].PatchAsync(renameItem);
					var renamedItem = await this.graphServiceClient.Drives[userDriveId].Items[data.FirstOrDefault().ParentId].Children[newName].GetAsync();
					file.Name = renamedItem.Name;
					file.Id = renamedItem.Id;
					file.ParentId = renamedItem.ParentReference?.Id;
					file.Size = renamedItem.Size ?? 0;
					file.IsFile = renamedItem.Folder == null ? true : false;
					file.DateModified = renamedItem.LastModifiedDateTime.Value.DateTime;
					file.DateCreated = renamedItem.CreatedDateTime.Value.DateTime;
					file.HasChild = renamedItem.Folder != null && renamedItem.Folder.ChildCount > 0 ? true : false;
					file.Type = Path.GetExtension(renamedItem.Name);
					file.FilterId = await GetFilterId(renamedItem);
					file.FilterPath = await GetFilterPath(renamedItem);
					file.Permission = GetPathPermission(file.FilterPath + (file.IsFile ? file.Name : Path.GetFileNameWithoutExtension(file.Name)), file.IsFile);
				}
				renameResponse.Files = new List<FileManagerDirectoryContent> { file };
				return renameResponse;
			}
			catch(Exception ex)
			{
				ErrorDetails er = new ErrorDetails();
				er.Message = (ex.GetType().Name == "UnauthorizedAccessException") ? "'" + name + "' is not accessible. You need permission to perform the write action." : ex.Message.ToString();
				er.Code = er.Message.Contains("is not accessible. You need permission") ? "401" : "417";
				if ((er.Code == "401") && !string.IsNullOrEmpty(accessMessage)) { er.Message = accessMessage; }
				renameResponse.Error = er;
				return renameResponse;
			}
		}

		public FileManagerResponse Upload(string path, IList<IFormFile> uploadFiles, string action, FileManagerDirectoryContent[] data)
		{
			return AsyncUpload(path, uploadFiles, action, data).Result;
		}

		public virtual async Task<FileManagerResponse> AsyncUpload(string path, IList<IFormFile> uploadFiles, string action, FileManagerDirectoryContent[] data)
		{
			if (string.IsNullOrEmpty(userDriveId))
			{
				await GetSharePointDrive();
			}
			FileManagerResponse uploadResponse = new FileManagerResponse();
			AccessPermission PathPermission = GetPathPermission(data[0].FilterPath + data[0].Name, false);
			try
			{
				if (PathPermission != null && (!PathPermission.Read || !PathPermission.Upload))
				{
					accessMessage = PathPermission.Message;
					throw new UnauthorizedAccessException("'" + data[0].Name + "' is not accessible. You need permission to perform the upload action.");
				}
				var driveItemId = data[0].Id;
				List<string> existFiles = new List<string>();

				foreach (IFormFile file in uploadFiles)
				{
					string fileName = Path.GetFileName(file.FileName);
					fileName = fileName.Replace("../", "");

					string[] folders = file.FileName.Split('/');
					if (folders.Length > 1)
					{
						string currentPath = path;
						string finalFolderName = folders[folders.Length - 2];
						string parentPath = driveItemId;

						for (int i = 0; i < folders.Length - 1; i++)
						{
							string folderName = i == folders.Length - 2 ? finalFolderName : folders[i];
							currentPath = Path.Combine(currentPath, folderName);

							if(await ItemExistsInDrive(driveItemId, currentPath))
							{
								if (action == "keepboth")
								{
									currentPath = GenerateUniqueFileName(driveItemId, currentPath);
								}
								else if (action == "remove")
								{
									var removeResult = await RemoveItemFromDriveAsync(driveItemId, fileName);
									if (!removeResult.Success)
									{
										ErrorDetails er = new ErrorDetails
										{
											Code = "404",
											Message = "File not found."
										};
										uploadResponse.Error = er;
									}
								}
							}
							var createData = new FileManagerDirectoryContent
							{
								Id = parentPath,
								Name = folderName,
								FilterPath = currentPath
							};

							await createAsync(currentPath, folderName, createData);

							var createdFolder = await GetDriveItemByName(driveItemId, currentPath);
							if (createdFolder != null)
							{
								parentPath = createdFolder.Id;
							}
						}

						string fullPath = Path.Combine(currentPath, folders.Last());
						var uploadResult = await UploadFileToDriveAsync(driveItemId, file, action, fullPath);

						if (!uploadResult.Success)
						{
							existFiles.Add(fileName);
						}
					}
					else
					{
						if (uploadFiles != null)
						{
							if (action == "save")
							{
								var uploadResult = await UploadFileToDriveAsync(driveItemId, file, action, null);
								if (!uploadResult.Success)
								{
									existFiles.Add(fileName);
								}
							}
							else if (action == "replace")
							{
								var uploadResult = await UploadFileToDriveAsync(driveItemId, file, action, null);
								if (!uploadResult.Success)
								{
									existFiles.Add(fileName);
								}
							}
							else if (action == "keepboth")
							{
								string newFileName = GenerateUniqueFileName(driveItemId, file.FileName);
								fileName = Path.Combine(Path.GetTempPath(), newFileName);

								var uploadResult = await UploadFileToDriveAsync(driveItemId, file, action, newFileName);
								if (!uploadResult.Success)
								{
									existFiles.Add(newFileName);
								}
							}
							else if (action == "remove")
							{
								var removeResult = await RemoveItemFromDriveAsync(driveItemId, fileName);
								if (!removeResult.Success)
								{
									ErrorDetails er = new ErrorDetails
									{
										Code = "404",
										Message = "File not found."
									};
									uploadResponse.Error = er;
								}
							}
						}
					}
				}

				if (existFiles.Count != 0)
				{
					ErrorDetails er = new ErrorDetails
					{
						FileExists = existFiles,
						Code = "400",
						Message = "File Already Exists"
					};
					uploadResponse.Error = er;
				}

				return uploadResponse;
			}
			catch (Exception ex)
			{
				ErrorDetails er = new ErrorDetails
				{
					Message = (ex.GetType().Name == "UnauthorizedAccessException") ? "'" + data[0].Name + "' is not accessible. You need permission to perform the upload action." : ex.Message.ToString(),
					Code = ex.GetType().Name == "UnauthorizedAccessException" ? "401" : "417"
				};
				if (er.Code == "401" && !string.IsNullOrEmpty(accessMessage)) { er.Message = accessMessage; }
				uploadResponse.Error = er;
				return uploadResponse;
			}
		}

		private async Task<(bool Success, string Message)> UploadFileToDriveAsync(string driveItemId, IFormFile file, string action, string newFileName = null)
		{
			try
			{
				if (action != "replace")
				{
					bool isExist = await ItemExistsInDrive(driveItemId, newFileName != null ? newFileName : file.FileName);
					if (isExist)
					{
						return (false, "File already exists");
					}
				}
				var fileName = newFileName ?? file.FileName;
				using (var fileStream = file.OpenReadStream())
				{
					var itemPath = $"{fileName}";
					await this.graphServiceClient.Drives[userDriveId].Items[driveItemId].ItemWithPath(itemPath).Content.PutAsync(fileStream);
				}
				return (true, "Upload successful");
			}
			catch (ServiceException ex)
			{
				return (false, ex.Message);
			}
		}

		private async Task<DriveItem> GetDriveItemByName(string driveItemId, string folderPath)
		{
			var items = await this.graphServiceClient.Drives[userDriveId].Items[driveItemId].Children.GetAsync();
			return items.Value.FirstOrDefault(item => item.Name.Equals(Path.GetFileName(folderPath), StringComparison.OrdinalIgnoreCase));
		}

		public async Task<bool> ItemExistsInDrive(string driveItemId, string itemName)
		{
			try
			{
				var item = await this.graphServiceClient.Drives[userDriveId].Items[driveItemId].ItemWithPath(itemName).GetAsync();
				return item != null;
			}
			catch (Exception)
			{
				return false;
			}
		}

		private async Task<(bool Success, string Message)> RemoveItemFromDriveAsync(string driveItemId, string itemName)
		{
			try
			{
				await this.graphServiceClient.Drives[userDriveId].Items[driveItemId].ItemWithPath(itemName).DeleteAsync();
				return (true, "Item deleted successfully");
			}
			catch (Exception)
			{
				return (false, "Item failed to delete");
			}
		}

		private string GenerateUniqueFileName(string driveItemId, string name)
		{
			try
			{
				string newName = name;
				int count = 0;
				string extension = Path.GetExtension(name);
				string nameWithoutExtension = Path.GetFileNameWithoutExtension(name);

				while (ItemExistsInDrive(driveItemId, newName).Result)
				{
					count++;
					newName = $"{nameWithoutExtension}({count}){extension}";
				}

				return newName;
			}
			catch (Exception ex)
			{
				return ex.Message;
			}
		}

		public virtual FileStreamResult Download(string path, string[] names, params FileManagerDirectoryContent[] data)
		{
			return DownloadSelectedDriveItemsAsync(path, names, data).Result;
		}

		public async Task<FileStreamResult> DownloadSelectedDriveItemsAsync(string path, string[] names, params FileManagerDirectoryContent[] data)
		{
			if (string.IsNullOrEmpty(userDriveId))
			{
				await GetSharePointDrive();
			}
			try
			{
				if (data.Length == 1 && data[0].IsFile)
				{
					AccessPermission pathPermission = GetPathPermission(data[0].FilterPath + (data[0].IsFile ? data[0].Name : Path.GetFileNameWithoutExtension(data[0].Name)), data[0].IsFile);
					if (pathPermission != null && (!pathPermission.Read || !pathPermission.Download))
					{
						throw new UnauthorizedAccessException("'" + data[0].Name + "' is not accessible. Access is denied.");
					}
					Stream stream = await this.graphServiceClient.Drives[userDriveId].Items[data[0].Id].Content.GetAsync();
					return new FileStreamResult(stream, "APPLICATION/octet-stream")
					{
						FileDownloadName = data[0].Name
					};
				}

				var memoryStream = new MemoryStream();
				using (var archive = new ZipArchive(memoryStream, ZipArchiveMode.Create, true))
				{
					foreach (var item in data)
					{
						AccessPermission pathPermission = GetPathPermission(item.FilterPath + (item.IsFile ? item.Name : Path.GetFileNameWithoutExtension(item.Name)), item.IsFile);
						if (pathPermission != null && (!pathPermission.Read || !pathPermission.Download))
						{
							throw new UnauthorizedAccessException("'" + item.Name + "' is not accessible. Access is denied.");
						}
						var driveItem = await this.graphServiceClient.Drives[userDriveId].Items[item.Id].GetAsync();

						if (driveItem.File != null)
						{
							var downloadStream = await this.graphServiceClient.Drives[userDriveId].Items[item.Id].Content.GetAsync();
							var zipEntry = archive.CreateEntry(driveItem.Name);
							using (var zipEntryStream = zipEntry.Open())
							{
								await downloadStream.CopyToAsync(zipEntryStream);
							}
						}
						else if (driveItem.Folder != null)
						{
							await AddFolderToZipAsync(archive, driveItem, driveItem.Name);
						}
					}
				}

				memoryStream.Seek(0, SeekOrigin.Begin);
				return new FileStreamResult(memoryStream, "APPLICATION/octet-stream")
				{
					FileDownloadName = "SelectedFiles.zip"
				};
			}
			catch (Exception)
			{
				throw;
			}
		}

		private async Task AddFolderToZipAsync(ZipArchive archive, DriveItem folderItem, string currentFolderPath)
		{
			// Create the folder entry in the ZIP archive
			var folderEntry = archive.CreateEntry(currentFolderPath + "/");

			var folderContents = await this.graphServiceClient.Drives[userDriveId].Items[folderItem.Id].Children.GetAsync();

			foreach (var item in folderContents.Value)
			{
				if (item.File != null)
				{
					var downloadStream = await this.graphServiceClient.Drives[userDriveId].Items[item.Id].Content.GetAsync();

					var zipEntry = archive.CreateEntry(Path.Combine(currentFolderPath, item.Name));
					using (var zipEntryStream = zipEntry.Open())
					{
						await downloadStream.CopyToAsync(zipEntryStream);
					}
				}
				else if (item.Folder != null)
				{
					await AddFolderToZipAsync(archive, item, Path.Combine(currentFolderPath, item.Name));
				}
			}
		}

		public virtual FileStreamResult GetImage(string path, string id, bool allowCompress, params FileManagerDirectoryContent[] data)
		{
			return this.GetImageAsync(path, id, allowCompress, data).Result;
		}

		public async Task<FileStreamResult> GetImageAsync(string path, string id, bool allowCompress, params FileManagerDirectoryContent[] data)
		{
			if (string.IsNullOrEmpty(userDriveId))
			{
				await GetSharePointDrive();
			}
			Stream stream = await this.graphServiceClient.Drives[userDriveId].Items[id].Content.GetAsync();
			return new FileStreamResult(stream, "APPLICATION/octet-stream");
		}

		public async Task<List<FileManagerDirectoryContent>> GetChild(string itemId)
		{
			List<FileManagerDirectoryContent> child = new List<FileManagerDirectoryContent>();
			var children = await this.graphServiceClient.Drives[userDriveId].Items[itemId].Children.GetAsync();
			foreach (var item in children.Value)
			{
				FileManagerDirectoryContent file = new FileManagerDirectoryContent();
				file.Name = item.Name;
				file.Id = item.Id;
				file.ParentId = item.ParentReference?.Id;
				file.Size = item.Size ?? 0;
				file.IsFile = item.Folder == null ? true : false;
				file.DateModified = item.LastModifiedDateTime.Value.DateTime;
				file.DateCreated = item.CreatedDateTime.Value.DateTime;
				file.HasChild = item.File != null ? false : await HasChild(item);
				file.Type = Path.GetExtension(item.Name);
				file.FilterId = await GetFilterId(item);
				file.FilterPath = await GetFilterPath(item);
				file.Permission = GetPathPermission(file.FilterPath + (file.IsFile ? file.Name : Path.GetFileNameWithoutExtension(file.Name)), file.IsFile);
				child.Add(file);
			}
			return child;
		}

		private String byteConversion(long fileSize)
		{
			try
			{
				string[] index = { "B", "KB", "MB", "GB", "TB", "PB", "EB" };
				if (fileSize == 0)
				{
					return "0 " + index[0];
				}

				long bytes = Math.Abs(fileSize);
				int loc = Convert.ToInt32(Math.Floor(Math.Log(bytes, 1024)));
				double num = Math.Round(bytes / Math.Pow(1024, loc), 1);
				return (Math.Sign(fileSize) * num).ToString() + " " + index[loc];
			}
			catch (Exception)
			{
				throw;
			}
		}

		private async Task<bool> HasChild(DriveItem item)
		{
			var children = await this.graphServiceClient.Drives[userDriveId].Items[item.Id].Children.GetAsync();
			if (children.Value.Any(child => child.Folder != null))
			{
				return true;
			}
			else
			{
				return false;
			}
		}

		private async Task<string> GetFilterId(DriveItem item)
		{
			if (item.ParentReference != null && item.ParentReference?.Id != null)
			{
				var currentItem = item;
				var parentIds = new List<string>();
				while (currentItem?.ParentReference != null && currentItem.ParentReference.Id != null)
				{
					parentIds.Insert(0, currentItem.ParentReference.Id);

					currentItem = await this.graphServiceClient.Drives[userDriveId].Items[currentItem.ParentReference.Id].GetAsync();
				}
				var filterID = string.Join("/", parentIds) + "/";

				return filterID;
			}
			else
			{
				return "";
			}
		}

		private async Task<string> GetFilterPath(DriveItem item)
		{
			if (item.ParentReference != null && item.ParentReference.Path != null)
			{
				var currentItem = item;
				var parentPaths = new List<string>();
				while (currentItem?.ParentReference != null && currentItem.ParentReference.Path != null)
				{
					var filterNames = currentItem.ParentReference.Path.Split(new[] { ":/" }, StringSplitOptions.RemoveEmptyEntries);

					if (filterNames.Length > 1)
					{
						var folderStructure = filterNames[1];

						var folderNames = folderStructure.Split(new[] { '/' }, StringSplitOptions.RemoveEmptyEntries);

						if (folderNames.Any())
						{
							var folderName = folderNames.Last();
							parentPaths.Insert(0, folderName);
						}
					}

					currentItem = await this.graphServiceClient.Drives[userDriveId].Items[currentItem.ParentReference.Id].GetAsync();
				}

				var filterPath = "/" + string.Join("/", parentPaths);
				if (parentPaths.Count > 0)
				{
					filterPath += "/";
				}

				return filterPath;
			}
			else
			{
				return "";
			}
		}


		protected virtual AccessPermission GetPathPermission(string path, bool isFile)
		{
			string[] fileDetails = GetFolderDetails(path);
			if (isFile)
			{
				return GetPermission(fileDetails[0].TrimStart('/') + fileDetails[1], fileDetails[1], true);
			}
			return GetPermission(fileDetails[0].TrimStart('/') + fileDetails[1], fileDetails[1], false);

		}

		protected virtual string[] GetFolderDetails(string path)
		{
			string[] str_array = path.Split('/'), fileDetails = new string[2];
			string parentPath = "";
			for (int i = 0; i < str_array.Length - 1; i++)
			{
				parentPath += str_array[i] + "/";
			}
			fileDetails[0] = parentPath;
			fileDetails[1] = str_array[str_array.Length - 1];
			return fileDetails;
		}

		protected virtual AccessPermission GetPermission(string location, string name, bool isFile)
		{
			AccessPermission permission = new AccessPermission();
			if (!isFile)
			{
				if (this.AccessDetails.AccessRules == null) { return null; }
				foreach (AccessRule folderRule in AccessDetails.AccessRules)
				{
					if (folderRule.Path != null && folderRule.IsFile == false && (folderRule.Role == null || folderRule.Role == AccessDetails.Role))
					{
						if (folderRule.Path.IndexOf("*") > -1)
						{
							string parentPath = folderRule.Path.Substring(0, folderRule.Path.IndexOf("*"));
							if ((location).IndexOf((parentPath)) == 0 || parentPath == "")
							{
								permission = UpdateFolderRules(permission, folderRule);
							}
						}
						else if ((folderRule.Path) == (location) || (folderRule.Path) == (location + Path.DirectorySeparatorChar) || (folderRule.Path) == (location + "/"))
						{
							permission = UpdateFolderRules(permission, folderRule);
						}
						else if ((location).IndexOf((folderRule.Path)) == 0)
						{
							permission = UpdateFolderRules(permission, folderRule);
						}
					}
				}
				return permission;
			}
			else
			{
				if (this.AccessDetails.AccessRules == null) return null;
				string nameExtension = Path.GetExtension(name).ToLower();
				string fileName = Path.GetFileNameWithoutExtension(name);
				string currentPath = (location + "/");
				foreach (AccessRule fileRule in AccessDetails.AccessRules)
				{
					if (!string.IsNullOrEmpty(fileRule.Path) && fileRule.IsFile && (fileRule.Role == null || fileRule.Role == AccessDetails.Role))
					{
						if (fileRule.Path.IndexOf("*.*") > -1)
						{
							string parentPath = fileRule.Path.Substring(0, fileRule.Path.IndexOf("*.*"));
							if (currentPath.IndexOf((parentPath)) == 0 || parentPath == "")
							{
								permission = UpdateFileRules(permission, fileRule);
							}
						}
						else if (fileRule.Path.IndexOf("*.") > -1)
						{
							string pathExtension = Path.GetExtension(fileRule.Path).ToLower();
							string parentPath = fileRule.Path.Substring(0, fileRule.Path.IndexOf("*."));
							if (((parentPath) == currentPath || parentPath == "") && nameExtension == pathExtension)
							{
								permission = UpdateFileRules(permission, fileRule);
							}
						}
						else if (fileRule.Path.IndexOf(".*") > -1)
						{
							string pathName = Path.GetFileNameWithoutExtension(fileRule.Path);
							string parentPath = fileRule.Path.Substring(0, fileRule.Path.IndexOf(pathName + ".*"));
							if (((parentPath) == currentPath || parentPath == "") && fileName == pathName)
							{
								permission = UpdateFileRules(permission, fileRule);
							}
						}
						else if ((fileRule.Path) == (Path.GetFileNameWithoutExtension(location)) || fileRule.Path == location || (fileRule.Path + nameExtension == location))
						{
							permission = UpdateFileRules(permission, fileRule);
						}
					}
				}
				return permission;
			}

		}

		protected virtual AccessPermission UpdateFolderRules(AccessPermission folderPermission, AccessRule folderRule)
		{
			folderPermission.Copy = HasPermission(folderRule.Copy);
			folderPermission.Download = HasPermission(folderRule.Download);
			folderPermission.Write = HasPermission(folderRule.Write);
			folderPermission.WriteContents = HasPermission(folderRule.WriteContents);
			folderPermission.Read = HasPermission(folderRule.Read);
			folderPermission.Upload = HasPermission(folderRule.Upload);
			folderPermission.Message = string.IsNullOrEmpty(folderRule.Message) ? string.Empty : folderRule.Message;
			return folderPermission;
		}

		protected virtual AccessPermission UpdateFileRules(AccessPermission filePermission, AccessRule fileRule)
		{
			filePermission.Copy = HasPermission(fileRule.Copy);
			filePermission.Download = HasPermission(fileRule.Download);
			filePermission.Write = HasPermission(fileRule.Write);
			filePermission.Read = HasPermission(fileRule.Read);
			filePermission.Message = string.IsNullOrEmpty(fileRule.Message) ? string.Empty : fileRule.Message;
			return filePermission;
		}

		protected virtual bool HasPermission(Syncfusion.Web.FileManager.Base.Permission rule)
		{
			return rule == Syncfusion.Web.FileManager.Base.Permission.Allow ? true : false;
		}

		public string ToCamelCase(FileManagerResponse userData)
		{
            return JsonConvert.SerializeObject(userData, new JsonSerializerSettings
            {
                ContractResolver = new DefaultContractResolver
                {
                    NamingStrategy = new CamelCaseNamingStrategy()
                }
            });
        }
	}
}
