using Newtonsoft.Json;
using Apps.MicrosoftSharePoint.Actions;
using Apps.MicrosoftSharePoint.Models.Identifiers;

namespace Tests.MicrosoftSharePoint; 

[TestClass]
public class DriveActionsTests :TestBase
{
    [TestMethod]
    public async Task FindFolder_IsSuccess()
    {
        var action = new DriveActions(InvocationContext,FileManager);
       
        var result = await action.FindFolderByName(
            new ParentFolderIdentifier { ParentFolderId= "01C7WXPSHVF2MLHRDQM5GJNQXIMR5A3QGW" },
            "Backup");
            Console.WriteLine($"Key: {result.Id}, Value: {result.Name}");

        Assert.IsNotNull(result);
    }

    [TestMethod]
    public async Task GetFileMetadata_IsSuccess()
    {
        // Arrange
        var action = new DriveActions(InvocationContext, FileManager);
        var fileId = "017O7UAG2K5K5ZS2DUTVDYLK4QKARACM54";

        // Act
        var result = await action.GetFileMetadataById(new FileIdentifier { FileId = fileId });

        // Assert
        Console.WriteLine(JsonConvert.SerializeObject(result, Formatting.Indented));
        Assert.IsNotNull(result);
    }
}
