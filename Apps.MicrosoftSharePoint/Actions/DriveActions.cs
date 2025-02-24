using System.Net.Mime;
using Apps.MicrosoftSharePoint.Dtos;
using Apps.MicrosoftSharePoint.Models.Identifiers;
using Apps.MicrosoftSharePoint.Models.Requests;
using Apps.MicrosoftSharePoint.Models.Responses;
using Blackbird.Applications.Sdk.Common;
using Blackbird.Applications.Sdk.Common.Actions;
using Blackbird.Applications.Sdk.Common.Authentication;
using Blackbird.Applications.Sdk.Common.Invocation;
using Apps.MicrosoftSharePoint.Extensions;
using Blackbird.Applications.SDK.Extensions.FileManagement.Interfaces;
using Blackbird.Applications.Sdk.Utils.Extensions.Files;
using RestSharp;
using Blackbird.Applications.Sdk.Common.Exceptions;

namespace Apps.MicrosoftSharePoint.Actions;

[ActionList]
public class DriveActions : BaseInvocable
{
    private readonly IEnumerable<AuthenticationCredentialsProvider> _authenticationCredentialsProviders;
    private readonly MicrosoftSharePointClient _client;
    private readonly IFileManagementClient _fileManagementClient;

    public DriveActions(InvocationContext invocationContext, IFileManagementClient fileManagementClient) 
        : base(invocationContext)
    {
        _authenticationCredentialsProviders = invocationContext.AuthenticationCredentialsProviders;
        _client = new MicrosoftSharePointClient(_authenticationCredentialsProviders);
        _fileManagementClient = fileManagementClient;
    }
    
    #region File actions

    [Action("Get file metadata", Description = "Retrieve the metadata for a file from site documents.")]
    public async Task<FileMetadataDto> GetFileMetadataById([ActionParameter] FileIdentifier fileIdentifier)
    {
        var request = new MicrosoftSharePointRequest($"/drive/items/{fileIdentifier.FileId}", Method.Get, 
            _authenticationCredentialsProviders);
        var fileMetadata = await _client.ExecuteWithHandling<FileMetadataDto>(request);
        return fileMetadata;
    }

    [Action("List changed files", Description = "List all files that have been created or modified during past hours. " +
                                                "If number of hours is not specified, files changed during past 24 " +
                                                "hours are listed.")]
    public async Task<ListFilesResponse> ListChangedFiles([ActionParameter] [Display("Hours")] int? hours)
    {
        var endpoint = "/drive/root/search(q='.')?$orderby=lastModifiedDateTime desc";
        var startDateTime = (DateTime.Now - TimeSpan.FromHours(hours ?? 24)).ToUniversalTime();
        var changedFiles = new List<FileMetadataDto>();
        int filesCount;
    
        do
        {
            var request = new MicrosoftSharePointRequest(endpoint, Method.Get, _authenticationCredentialsProviders);
            var result = await _client.ExecuteWithHandling<ListWrapper<FileMetadataDto>>(request);
            var files = result.Value.Where(item => item.MimeType != null && item.LastModifiedDateTime >= startDateTime);
            filesCount = files.Count();
            changedFiles.AddRange(files);
            endpoint = result.ODataNextLink == null ? null : "/drive" + result.ODataNextLink?.Split("drive")[^1];
        } while (endpoint != null && filesCount != 0);
    
        return new ListFilesResponse { Files = changedFiles };
    }
    
    [Action("Download file", Description = "Download a file from site documents.")]
    public async Task<FileResponse> DownloadFileById([ActionParameter] FileIdentifier fileIdentifier)
    {
        var request = new MicrosoftSharePointRequest($"/drive/items/{fileIdentifier.FileId}/content", Method.Get, 
            _authenticationCredentialsProviders);
        var response = await _client.ExecuteWithHandling(request);
        var filename = response.ContentHeaders.First(h => h.Name == "Content-Disposition").Value.ToString().Split('"')[1];
        var contentType = response.ContentType == MediaTypeNames.Text.Plain
            ? MediaTypeNames.Text.RichText
            : response.ContentType;

        using var stream = new MemoryStream(response.RawBytes);
        var file = await _fileManagementClient.UploadAsync(stream, contentType, filename);
        return new FileResponse { File = file };
    }
    
    [Action("Upload file to folder", Description = "Upload a file to a parent folder.")]
    public async Task<FileMetadataDto> UploadFileInFolderById([ActionParameter] ParentFolderIdentifier folderIdentifier,
        [ActionParameter] UploadFileRequest input)
    {
        if (folderIdentifier.ParentFolderId.StartsWith("/"))
        {
            throw new PluginApplicationException("Incorrect parent folder ID. Please provide a valid folder ID, such as '01C7WXPSBPDJQQ2E2CTBFI5XXXXXXXXXX'.");
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
            var uploadRequest = new MicrosoftSharePointRequest($".//drive/items/{folderIdentifier.ParentFolderId}:/{input.File.Name}:" +
                                                               $"/content?@microsoft.graph.conflictBehavior={input.ConflictBehavior}",
                Method.Put, _authenticationCredentialsProviders);
            uploadRequest.AddParameter(contentType, fileBytes, ParameterType.RequestBody);
            fileMetadata = await _client.ExecuteWithHandling<FileMetadataDto>(uploadRequest);
        }
        else
        {
            const int chunkSize = 3932160;
    
            var createUploadSessionRequest = new MicrosoftSharePointRequest(
                $".//drive/items/{folderIdentifier.ParentFolderId}:/{input.File.Name}:/createUploadSession", Method.Post,
                _authenticationCredentialsProviders);
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
            var uploadClient = new MicrosoftSharePointRestClient(baseUrl);

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
    
        return fileMetadata;
    }
    
    [Action("Delete file", Description = "Delete file from site documents.")]
    public async Task DeleteFileById([ActionParameter] FileIdentifier fileIdentifier)
    {
        var request = new MicrosoftSharePointRequest($"/drive/items/{fileIdentifier.FileId}", Method.Delete, 
            _authenticationCredentialsProviders); 
        await _client.ExecuteWithHandling(request);
    }
    
    #endregion
    
    #region Folder actions
    
    [Action("Get folder metadata", Description = "Retrieve the metadata for a folder.")]
    public async Task<FolderMetadataDto> GetFolderMetadataById([ActionParameter] FolderIdentifier folderIdentifier)
    {
        var request = new MicrosoftSharePointRequest($"/drive/items/{folderIdentifier.FolderId}", Method.Get, 
            _authenticationCredentialsProviders);
        var folderMetadata = await _client.ExecuteWithHandling<FolderMetadataDto>(request);
        return folderMetadata;
    }

    [Action("List files in folder", Description = "Retrieve metadata for files contained in a folder.")]
    public async Task<ListFilesResponse> ListFilesInFolderById([ActionParameter] FolderIdentifier folderIdentifier)
    {
        var filesInFolder = new List<FileMetadataDto>();
        var endpoint = $"/drive/items/{folderIdentifier.FolderId}/children";
        
        do
        {
            var request = new MicrosoftSharePointRequest(endpoint, Method.Get, _authenticationCredentialsProviders);
            var result = await _client.ExecuteWithHandling<ListWrapper<FileMetadataDto>>(request);
            var files = result.Value.Where(item => item.MimeType != null);
            filesInFolder.AddRange(files);
            endpoint = result.ODataNextLink == null ? null : "/drive" + result.ODataNextLink?.Split("drive")[^1];
        } while (endpoint != null);
        
        return new ListFilesResponse { Files = filesInFolder };
    }
    
    [Action("Create folder in parent folder", Description = "Create a new folder in parent folder.")]
    public async Task<FolderMetadataDto> CreateFolderInParentFolderWithId(
        [ActionParameter] ParentFolderIdentifier folderIdentifier,
        [ActionParameter] [Display("Folder name")] string folderName)
    {
        var request = new MicrosoftSharePointRequest($"/drive/items/{folderIdentifier.ParentFolderId}/children", 
            Method.Post, _authenticationCredentialsProviders);
        request.AddJsonBody(new
        {
            Name = folderName.Trim(),
            Folder = new { }
        });
    
        var folderMetadata = await _client.ExecuteWithHandling<FolderMetadataDto>(request);
        return folderMetadata;
    }
    
    [Action("Delete folder", Description = "Delete a folder.")]
    public async Task DeleteFolderById([ActionParameter] FolderIdentifier folderIdentifier)
    {
        var request = new MicrosoftSharePointRequest($"/drive/items/{folderIdentifier.FolderId}", Method.Delete, 
            _authenticationCredentialsProviders); 
        await _client.ExecuteWithHandling(request);
    }

    #endregion
}