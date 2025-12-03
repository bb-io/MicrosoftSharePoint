using RestSharp;
using System.Net.Mime;
using Apps.MicrosoftSharePoint.Dtos;
using Apps.MicrosoftSharePoint.Helper;
using Apps.MicrosoftSharePoint.Extensions;
using Apps.MicrosoftSharePoint.Models.Requests;
using Apps.MicrosoftSharePoint.Models.Entities;
using Apps.MicrosoftSharePoint.Models.Responses;
using Apps.MicrosoftSharePoint.Models.Identifiers;
using Blackbird.Applications.Sdk.Common;
using Blackbird.Applications.SDK.Blueprints;
using Blackbird.Applications.Sdk.Common.Actions;
using Blackbird.Applications.Sdk.Common.Invocation;
using Blackbird.Applications.Sdk.Common.Exceptions;
using Blackbird.Applications.Sdk.Common.Authentication;
using Blackbird.Applications.Sdk.Utils.Extensions.Files;
using Blackbird.Applications.SDK.Extensions.FileManagement.Interfaces;

namespace Apps.MicrosoftSharePoint.Actions;

[ActionList("Files")]
public class DriveActions : BaseInvocable
{
    private readonly IEnumerable<AuthenticationCredentialsProvider> _authenticationCredentialsProviders;
    private readonly SharePointBetaClient _client;
    private readonly IFileManagementClient _fileManagementClient;

    public DriveActions(InvocationContext invocationContext, IFileManagementClient fileManagementClient) 
        : base(invocationContext)
    {
        _authenticationCredentialsProviders = invocationContext.AuthenticationCredentialsProviders;
        _client = new SharePointBetaClient(_authenticationCredentialsProviders);
        _fileManagementClient = fileManagementClient;
    }
    
    #region File actions

    [Action("Get file metadata", Description = "Retrieve the metadata for a file from site documents.")]
    public async Task<FileMetadataDto> GetFileMetadataById([ActionParameter] FileIdentifier fileIdentifier)
    {
        var location = ItemIdParser.Parse(fileIdentifier.FileId);

        string endpoint;
        if (location.IsDefaultDrive)
            endpoint = $"/drive/items/{location.ItemId}?expand=listItem";
        else
            endpoint = $"/drives/{location.DriveId}/items/{location.ItemId}?expand=listItem";

        var request = new SharePointRequest(endpoint, Method.Get, _authenticationCredentialsProviders);
        var fileMetadata = await _client.ExecuteWithHandling<FileMetadataDto>(request);

        DriveEntity defaultDrive = await GetDefaultDrive();
        ProcessFileMetadataIds(fileMetadata, location, defaultDrive);   
        return fileMetadata;
    }

    [Action("List changed files", Description = "List all files that have been created or modified during past hours. " +
                                                "If number of hours is not specified, files changed during past 24 " +
                                                "hours are listed.")]
    public async Task<ListFilesResponse> ListChangedFiles(
        [ActionParameter] FolderIdentifier folderIdentifier,
        [ActionParameter][Display("Hours")] int? hours)
    {
        var location = ItemIdParser.Parse(folderIdentifier?.FolderId);

        string baseEndpoint;
        if (location.IsDefaultDrive)
        {
            baseEndpoint = location.ItemId.Equals("root", StringComparison.OrdinalIgnoreCase)
                ? "/drive/root"
                : $"/drive/items/{location.ItemId}";
        }
        else
        {
            baseEndpoint = location.ItemId.Equals("root", StringComparison.OrdinalIgnoreCase)
                ? $"/drives/{location.DriveId}/root"
                : $"/drives/{location.DriveId}/items/{location.ItemId}";
        }

        var endpoint = $"{baseEndpoint}/search(q='.')?$orderby=lastModifiedDateTime desc";

        var startDateTime = (DateTime.Now - TimeSpan.FromHours(hours ?? 24)).ToUniversalTime();
        var changedFiles = new List<FileMetadataDto>();
        int filesCount;

        do
        {
            var request = Uri.IsWellFormedUriString(endpoint, UriKind.Absolute)
                ? new SharePointRequest(new Uri(endpoint).ToString(), Method.Get, _authenticationCredentialsProviders)
                : new SharePointRequest(endpoint, Method.Get, _authenticationCredentialsProviders);

            var result = await _client.ExecuteWithHandling<ListWrapper<FileMetadataDto>>(request);

            if (result?.Value != null)
            {
                var files = result.Value
                    .Where(item => item.MimeType != null && item.LastModifiedDateTime >= startDateTime);

                filesCount = files.Count();
                changedFiles.AddRange(files);

                endpoint = result.ODataNextLink;
            }
            else
            {
                filesCount = 0;
                endpoint = null;
            }

        } while (!string.IsNullOrEmpty(endpoint) && filesCount != 0);

        DriveEntity defaultDrive = await GetDefaultDrive();
        foreach (var file in changedFiles)
            ProcessFileMetadataIds(file, location, defaultDrive);

        return new ListFilesResponse { Files = changedFiles };
    }

    [BlueprintActionDefinition(BlueprintAction.DownloadFile)]
    [Action("Download file", Description = "Download a file from site documents.")]
    public async Task<FileResponse> DownloadFileById([ActionParameter] FileIdentifier fileIdentifier)
    {
        var location = ItemIdParser.Parse(fileIdentifier.FileId);

        string endpoint;
        if (location.IsDefaultDrive)
            endpoint = $"/drive/items/{location.ItemId}/content";
        else
            endpoint = $"/drives/{location.DriveId}/items/{location.ItemId}/content";

        var request = new SharePointRequest(endpoint, Method.Get, _authenticationCredentialsProviders);
        var response = await _client.ExecuteWithHandling(request);

        var filename = response.ContentHeaders.First(h => h.Name == "Content-Disposition").Value.ToString().Split('"')[1];
        var contentType = response.ContentType ?? MediaTypeNames.Text.Plain;

        using var stream = new MemoryStream(response.RawBytes);
        var file = await _fileManagementClient.UploadAsync(stream, contentType, filename);

        return new FileResponse { File = file };
    }

    [BlueprintActionDefinition(BlueprintAction.UploadFile)]
    [Action("Upload file", Description = "Upload a file to a parent folder.")]
    public async Task<FileMetadataDto> UploadFileInFolderById(
        [ActionParameter] ParentFolderIdentifier folderIdentifier,
        [ActionParameter] UploadFileRequest input)
    {
        if (string.IsNullOrEmpty(folderIdentifier.ParentFolderId))
            throw new PluginMisconfigurationException("Parent folder ID cannot be empty.");

        if (input.File == null || input.File?.Name == null)
            throw new PluginMisconfigurationException("File is null. Please provide a valid file.");

        var location = ItemIdParser.Parse(folderIdentifier.ParentFolderId);

        string baseEndpoint;
        if (location.IsDefaultDrive)
        {
            baseEndpoint = location.ItemId.Equals("root", StringComparison.OrdinalIgnoreCase)
                ? "/drive/root"
                : $"/drive/items/{location.ItemId}";
        }
        else
        {
            baseEndpoint = location.ItemId.Equals("root", StringComparison.OrdinalIgnoreCase)
                ? $"/drives/{location.DriveId}/root"
                : $"/drives/{location.DriveId}/items/{location.ItemId}";
        }

        const int fourMegabytesInBytes = 4194304;
        var file = await _fileManagementClient.DownloadAsync(input.File);
        var fileBytes = await file.GetByteData();
        var fileSize = fileBytes.LongLength;
        var contentType = Path.GetExtension(input.File.Name) == ".txt"
            ? MediaTypeNames.Text.Plain
            : input.File.ContentType;
        var fileMetadata = new FileMetadataDto();

        if (fileSize < fourMegabytesInBytes)
        {
            var url = $"{baseEndpoint}:/{input.File.Name}:/content?@microsoft.graph.conflictBehavior={input.ConflictBehavior}";

            var uploadRequest = new SharePointRequest(url, Method.Put, _authenticationCredentialsProviders);
            uploadRequest.AddParameter(contentType, fileBytes, ParameterType.RequestBody);

            fileMetadata = await _client.ExecuteWithHandling<FileMetadataDto>(uploadRequest);
        }
        else
        {
            const int chunkSize = 3932160;

            var createSessionUrl = $"{baseEndpoint}:/{input.File.Name}:/createUploadSession";
            var createUploadSessionRequest = new SharePointRequest(
                createSessionUrl, 
                Method.Post, 
                _authenticationCredentialsProviders
            );

            createUploadSessionRequest.AddJsonBody($@"
            {{
                ""deferCommit"": false,
                ""item"": {{
                    ""@microsoft.graph.conflictBehavior"": ""{input.ConflictBehavior}"",
                    ""name"": ""{input.File.Name}""
                }}
            }}");

            var resumableUploadResult = await _client.ExecuteWithHandling<ResumableUploadDto>(createUploadSessionRequest);
            var uploadUrl = new Uri(resumableUploadResult.UploadUrl);
            var baseUrl = uploadUrl.GetLeftPart(UriPartial.Authority);
            var endpoint = uploadUrl.PathAndQuery;
            var uploadClient = new SharePointClient(baseUrl);

            do
            {
                var startByte = int.Parse(resumableUploadResult.NextExpectedRanges.First().Split("-")[0]);
                var buffer = fileBytes.Skip(startByte).Take(chunkSize).ToArray();
                var bufferSize = buffer.Length;

                var uploadRequest = new RestRequest(endpoint, Method.Put);
                uploadRequest.AddParameter(contentType, buffer, ParameterType.RequestBody);
                uploadRequest.AddHeader("Content-Length", bufferSize);
                uploadRequest.AddHeader("Content-Range", $"bytes {startByte}-{startByte + bufferSize - 1}/{fileSize}");

                var uploadResponse = await uploadClient.ExecuteWithHandling(uploadRequest);
                var responseContent = uploadResponse.Content;

                resumableUploadResult = responseContent.DeserializeObject<ResumableUploadDto>();

                if (resumableUploadResult.NextExpectedRanges == null)
                    fileMetadata = responseContent.DeserializeObject<FileMetadataDto>();

            } while (resumableUploadResult.NextExpectedRanges != null);
        }

        DriveEntity defaultDrive = await GetDefaultDrive();
        ProcessFileMetadataIds(fileMetadata, location, defaultDrive);
        return fileMetadata;
    }

    [Action("Delete file", Description = "Delete file from site documents.")]
    public async Task DeleteFileById([ActionParameter] FileIdentifier fileIdentifier)
    {
        if (string.IsNullOrEmpty(fileIdentifier.FileId))
            throw new PluginMisconfigurationException("File ID cannot be empty. Please provide a valid file ID.");

        try
        {
            await GetFileMetadataById(fileIdentifier);
        }
        catch (Exception ex)
        {
            throw new PluginApplicationException(
                $"Failed to verify file existence for ID {fileIdentifier.FileId}. Error: {ex.Message}"
            );
        }

        var location = ItemIdParser.Parse(fileIdentifier.FileId);

        string endpoint;
        if (location.IsDefaultDrive)
            endpoint = $"/drive/items/{location.ItemId}";
        else
            endpoint = $"/drives/{location.DriveId}/items/{location.ItemId}";

        var request = new SharePointRequest(endpoint, Method.Delete, _authenticationCredentialsProviders);
        await _client.ExecuteWithHandling(request);
    }

    #endregion

    #region Folder actions

    [Action("Get folder metadata", Description = "Retrieve the metadata for a folder.")]
    public async Task<FolderMetadataDto> GetFolderMetadataById([ActionParameter] FolderIdentifier folderIdentifier)
    {
        var location = ItemIdParser.Parse(folderIdentifier.FolderId);

        string endpoint;
        if (location.IsDefaultDrive)
        {
            endpoint = location.ItemId.Equals("root", StringComparison.OrdinalIgnoreCase)
                ? "/drive/root"
                : $"/drive/items/{location.ItemId}";
        }
        else
        {
            endpoint = location.ItemId.Equals("root", StringComparison.OrdinalIgnoreCase)
                ? $"/drives/{location.DriveId}/root"
                : $"/drives/{location.DriveId}/items/{location.ItemId}";
        }

        var request = new SharePointRequest(endpoint, Method.Get, _authenticationCredentialsProviders);
        var folderMetadata = await _client.ExecuteWithHandling<FolderMetadataDto>(request);

        DriveEntity defaultDrive = await GetDefaultDrive();
        ProcessFolderMetadataIds(folderMetadata, location, defaultDrive);
        return folderMetadata;
    }

    [Action("Search files", Description = "Retrieve metadata for files contained in a folder.")]
    public async Task<ListFilesResponse> ListFilesInFolderById(
        [ActionParameter] FolderIdentifier folderIdentifier,
        [ActionParameter] FilterExtensions extensions)
    {
        var location = ItemIdParser.Parse(folderIdentifier.FolderId);

        string endpoint;
        if (location.IsDefaultDrive)
        {
            endpoint = location.ItemId.Equals("root", StringComparison.OrdinalIgnoreCase)
                ? "/drive/root/children"
                : $"/drive/items/{location.ItemId}/children";
        }
        else
        {
            endpoint = location.ItemId.Equals("root", StringComparison.OrdinalIgnoreCase)
                ? $"/drives/{location.DriveId}/root/children"
                : $"/drives/{location.DriveId}/items/{location.ItemId}/children";
        }

        var filesInFolder = new List<FileMetadataDto>();
        do
        {
            var request = Uri.IsWellFormedUriString(endpoint, UriKind.Absolute)
                ? new SharePointRequest(new Uri(endpoint).ToString(), Method.Get, _authenticationCredentialsProviders)
                : new SharePointRequest(endpoint, Method.Get, _authenticationCredentialsProviders);

            var result = await _client.ExecuteWithHandling<ListWrapper<FileMetadataDto>>(request);

            if (result?.Value != null)
            {
                var files = result.Value.Where(item => item.MimeType != null);

                if (extensions != null && extensions.Extensions?.Any() == true)
                {
                    files = files.Where(item =>
                        !string.IsNullOrEmpty(item.Name) &&
                        extensions.Extensions.Any(ext =>
                            item.Name.EndsWith(ext, StringComparison.OrdinalIgnoreCase)));
                }

                filesInFolder.AddRange(files);
            }

            endpoint = result?.ODataNextLink;

        } while (!string.IsNullOrEmpty(endpoint));

        DriveEntity defaultDrive = await GetDefaultDrive();
        foreach (var file in filesInFolder)
            ProcessFileMetadataIds(file, location, defaultDrive);

        return new ListFilesResponse { Files = filesInFolder };
    }

    [Action("Create folder", Description = "Create a new folder in parent folder.")]
    public async Task<FolderMetadataDto> CreateFolderInParentFolderWithId(
        [ActionParameter] ParentFolderIdentifier folderIdentifier,
        [ActionParameter][Display("Folder name")] string folderName)
    {
        var location = ItemIdParser.Parse(folderIdentifier.ParentFolderId);

        string endpoint;
        if (location.IsDefaultDrive)
        {
            endpoint = location.ItemId.Equals("root", StringComparison.OrdinalIgnoreCase)
                ? "/drive/root/children"
                : $"/drive/items/{location.ItemId}/children";
        }
        else
        {
            endpoint = location.ItemId.Equals("root", StringComparison.OrdinalIgnoreCase)
                ? $"/drives/{location.DriveId}/root/children"
                : $"/drives/{location.DriveId}/items/{location.ItemId}/children";
        }

        var request = new SharePointRequest(endpoint, Method.Post, _authenticationCredentialsProviders);
        request.AddJsonBody(new
        {
            Name = folderName.Trim(),
            Folder = new { }
        });

        var folderMetadata = await _client.ExecuteWithHandling<FolderMetadataDto>(request);

        DriveEntity defaultDrive = await GetDefaultDrive();
        ProcessFolderMetadataIds(folderMetadata, location, defaultDrive);
        return folderMetadata;
    }

    [Action("Delete folder", Description = "Delete a folder.")]
    public async Task DeleteFolderById([ActionParameter] FolderIdentifier folderIdentifier)
    {
        if (string.IsNullOrEmpty(folderIdentifier.FolderId))
            throw new PluginMisconfigurationException("Folder ID cannot be empty.");

        var location = ItemIdParser.Parse(folderIdentifier.FolderId);

        if (location.ItemId.Equals("root", StringComparison.OrdinalIgnoreCase))
            throw new PluginMisconfigurationException("You cannot delete the root folder of a drive.");

        string endpoint;
        if (location.IsDefaultDrive)
            endpoint = $"/drive/items/{location.ItemId}";
        else
            endpoint = $"/drives/{location.DriveId}/items/{location.ItemId}";

        var request = new SharePointRequest(endpoint, Method.Delete, _authenticationCredentialsProviders);
        await _client.ExecuteWithHandling(request);
    }

    [Action("Find folder", Description = "Find a folder by name within a specified parent folder. Returns empty if not found.")]
    public async Task<FolderMetadataDto> FindFolderByName(
        [ActionParameter] ParentFolderIdentifier parentFolderIdentifier,
        [ActionParameter][Display("Folder name")] string folderName)
    {
        if (string.IsNullOrWhiteSpace(folderName))
            throw new PluginMisconfigurationException("Folder name cannot be empty.");

        var location = ItemIdParser.Parse(parentFolderIdentifier.ParentFolderId);

        string baseEndpoint;
        if (location.IsDefaultDrive)
        {
            baseEndpoint = location.ItemId.Equals("root", StringComparison.OrdinalIgnoreCase)
                ? "/drive/root"
                : $"/drive/items/{location.ItemId}";
        }
        else
        {
            baseEndpoint = location.ItemId.Equals("root", StringComparison.OrdinalIgnoreCase)
                ? $"/drives/{location.DriveId}/root"
                : $"/drives/{location.DriveId}/items/{location.ItemId}";
        }

        var endpoint = $"{baseEndpoint}/children?$select=id,name,folder";
        var folderNameTrimmed = folderName.Trim();

        try
        {
            DriveEntity defaultDrive = await GetDefaultDrive();
            do
            {
                var request = Uri.IsWellFormedUriString(endpoint, UriKind.Absolute)
                    ? new SharePointRequest(new Uri(endpoint).ToString(), Method.Get, _authenticationCredentialsProviders)
                    : new SharePointRequest(endpoint, Method.Get, _authenticationCredentialsProviders);

                var result = await _client.ExecuteWithHandling<ListWrapper<FolderMetadataDto>>(request);

                var folder = result.Value
                    .Where(item => item.ChildCount != null && item.Name.Contains(folderNameTrimmed, StringComparison.OrdinalIgnoreCase))
                    .FirstOrDefault();

                if (folder != null)
                {
                    ProcessFolderMetadataIds(folder, location, defaultDrive);
                    return folder;
                }

                endpoint = result.ODataNextLink;

            } while (!string.IsNullOrEmpty(endpoint));

            return new FolderMetadataDto();
        }
        catch (Exception ex)
        {
            throw new PluginApplicationException(
                $"Failed to find folder '{folderName}' in parent folder '{parentFolderIdentifier.ParentFolderId}'. " +
                $"Error: {ex.Message}"
            );
        }
    }

    #endregion

    private static FileMetadataDto ProcessFileMetadataIds(
        FileMetadataDto fileMetadata, 
        ItemLocationDto location, 
        DriveEntity defaultDrive)
    {
        var currentDriveId = location.DriveId ?? defaultDrive.Id;

        fileMetadata.FileId = ItemIdParser.Format(currentDriveId, fileMetadata.FileId, defaultDrive.Id);
        if (fileMetadata.ParentReference != null && !string.IsNullOrEmpty(fileMetadata.ParentReference.Id))
        {
            fileMetadata.ParentReference.Id = ItemIdParser.Format(
                currentDriveId, 
                fileMetadata.ParentReference.Id, 
                defaultDrive.Id
            );
        }
        return fileMetadata;
    }

    private static FolderMetadataDto ProcessFolderMetadataIds(
        FolderMetadataDto folderMetadata,
        ItemLocationDto location,
        DriveEntity defaultDrive)
    {
        var currentDriveId = location.DriveId ?? defaultDrive.Id;
        folderMetadata.Id = ItemIdParser.Format(currentDriveId, folderMetadata.Id, defaultDrive.Id);
        if (folderMetadata.ParentReference != null && !string.IsNullOrEmpty(folderMetadata.ParentReference.Id))
        {
            folderMetadata.ParentReference.Id = ItemIdParser.Format(
                currentDriveId,
                folderMetadata.ParentReference.Id,
                defaultDrive.Id
            );
        }
        return folderMetadata;
    }

    private async Task<DriveEntity> GetDefaultDrive()
    {
        var creds = InvocationContext.AuthenticationCredentialsProviders;
        var siteId = creds.First(x => x.KeyName == "SiteId").Value;

        var request = new SharePointRequest($"/sites/{siteId}/drive", Method.Get, creds);
        return await _client.ExecuteWithHandling<DriveEntity>(request);
    }
}