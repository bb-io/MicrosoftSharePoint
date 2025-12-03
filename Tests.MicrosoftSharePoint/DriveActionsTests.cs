using Tests.MicrosoftSharePoint.Base;
using Apps.MicrosoftSharePoint.Actions;
using Apps.MicrosoftSharePoint.Models.Requests;
using Apps.MicrosoftSharePoint.Models.Identifiers;
using Blackbird.Applications.Sdk.Common.Files;

namespace Tests.MicrosoftSharePoint; 

[TestClass]
public class DriveActionsTests : TestBase
{
    [TestMethod]
    public async Task FindFolder_IsSuccess()
    {
        // Arrange
        var action = new DriveActions(InvocationContext,FileManager);
        var folder = new ParentFolderIdentifier { ParentFolderId = "017O7UAG72P5YFPQ2R4NC3DPCGIRZPR54Y" };

        // Act
        var result = await action.FindFolderByName(folder, "tests");

        // Assert
        PrintJsonResult(result);
        Assert.IsNotNull(result);
    }

    [TestMethod]
    public async Task GetFileMetadata_IsSuccess()
    {
        // Arrange
        var action = new DriveActions(InvocationContext, FileManager);
        var fileId = "017O7UAG3PXHBWICU5ANAY55KTCPCF6TOJ";

        // Act
        var result = await action.GetFileMetadataById(new FileIdentifier { FileId = fileId });

        // Assert
        PrintJsonResult(result);
        Assert.IsNotNull(result);
    }

    [TestMethod]
    public async Task ListFilesInFolderById_IsSuccess()
    {
        // Arrange
        var action = new DriveActions(InvocationContext, FileManager);
        var folderId = "017O7UAG72P5YFPQ2R4NC3DPCGIRZPR54Y";

        // Act
        var result = await action.ListFilesInFolderById(
            new FolderIdentifier { FolderId = folderId },
            new FilterExtensions { Extensions = [".doc", ".docx"] }
        );

        // Assert
        PrintJsonResult(result);
        Assert.IsNotNull(result);
    }

    [TestMethod]
    public async Task GetFolderMetadataById_IsSuccess()
    {
        // Arrange
        var action = new DriveActions(InvocationContext, FileManager);
        var folder = new FolderIdentifier { FolderId = "b!V1tgT5LcyEiu-qCrLWi_sYE68iUDk7hCsFdpT5k21zfKEJWmu4ZAR7Zk_O7bubtF#017O7UAG7AGGYZEMSZFNDL6GYZFR6R2JGX" };

        // Act
        var result = await action.GetFolderMetadataById(folder);

        // Assert
        PrintJsonResult(result);
        Assert.IsNotNull(result);
    }

    [TestMethod]
    public async Task DownloadFileById_IsSuccess()
    {
        // Arrange
        var action = new DriveActions(InvocationContext, FileManager);
        var fileId = "017O7UAG3PXHBWICU5ANAY55KTCPCF6TOJ";

        // Act
        var result = await action.DownloadFileById(
            new FileIdentifier { FileId = fileId }
        );

        // Assert
        PrintJsonResult(result);
        Assert.IsNotNull(result);
    }

    [TestMethod]
    public async Task UploadFileInFolderById_IsSuccess()
    {
        // Arrange
        var action = new DriveActions(InvocationContext, FileManager);
        var folder = new ParentFolderIdentifier { ParentFolderId = "017O7UAG72P5YFPQ2R4NC3DPCGIRZPR54Y" };
        var input = new UploadFileRequest 
        { 
            File = new FileReference { Name = "uploaded.txt" },
            ConflictBehavior = "replace"
        };

        // Act
        var result = await action.UploadFileInFolderById(folder, input);

        // Assert
        PrintJsonResult(result);
        Assert.IsNotNull(result);
    }

    [TestMethod]
    public async Task DeleteFileById_IsSuccess()
    {
        // Arrange
        var action = new DriveActions(InvocationContext, FileManager);
        var fileId = "b!V1tgT5LcyEiu-qCrLWi_sYE68iUDk7hCsFdpT5k21zfKEJWmu4ZAR7Zk_O7bubtF#017O7UAG7HZXVAQYLD55B26A2A5HTSBLUE";

        // Act
        await action.DeleteFileById(new FileIdentifier { FileId = fileId });
    }

    [TestMethod]
    public async Task CreateFolderInParentFolderWithId_IsSuccess()
    {
        // Arrange
        var action = new DriveActions(InvocationContext, FileManager);
        var folder = new ParentFolderIdentifier { ParentFolderId = "b!V1tgT5LcyEiu-qCrLWi_sYE68iUDk7hCsFdpT5k21zfKEJWmu4ZAR7Zk_O7bubtF#017O7UAG7AGGYZEMSZFNDL6GYZFR6R2JGX" };

        // Act
        var result = await action.CreateFolderInParentFolderWithId(folder, "created from tests");

        // Assert
        PrintJsonResult(result);
        Assert.IsNotNull(result);
    }

    [TestMethod]
    public async Task DeleteFolderById_IsSuccess()
    {
        // Arrange
        var action = new DriveActions(InvocationContext, FileManager);
        var folderId = "017O7UAGZZGRGSJL7ILJEIVW6RARS2RH7M";

        // Act
        await action.DeleteFolderById(new FolderIdentifier { FolderId = folderId });
    }

    [TestMethod]
    public async Task ListChangedFiles_IsSuccess()
    {
        // Arrange
        var action = new DriveActions(InvocationContext, FileManager);
        var folder = new FolderIdentifier { FolderId = "b!V1tgT5LcyEiu-qCrLWi_sYE68iUDk7hCsFdpT5k21zfKEJWmu4ZAR7Zk_O7bubtF#017O7UAGZDLPJ5EDDMMBAKEEMBLWDFFBCC" };
        int hours = 1;

        // Act
        var result = await action.ListChangedFiles(folder, hours);

        // Assert
        PrintJsonResult(result);
        Assert.IsNotNull(result);
    }
}
