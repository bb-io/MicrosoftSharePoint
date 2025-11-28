using Apps.MicrosoftSharePoint.DataSourceHandlers;
using Apps.MicrosoftSharePoint.Models.Identifiers;
using Blackbird.Applications.Sdk.Common.Dynamic;
using Blackbird.Applications.SDK.Extensions.FileManagement.Models.FileDataSourceItems;

namespace Tests.MicrosoftSharePoint;

[TestClass]
public class DataHandlerTests : TestBase
{
    [TestMethod]
    public async Task FilePickerDataHandler_IsSuccess()
    {
        // Arrange
        var fileRequest = new FileIdentifier { DriveId = "b!vxtzBEWI9Eq4Vkw6hbeSgU7zvg6s6hJLqLwvmwOdy5MXAdvPapxlQKKjxaj1bN1o" };
        var handler = new FilePickerDataSourceHandler(InvocationContext, fileRequest);

        // Act
        var result = await handler.GetFolderContentAsync(
            new FolderContentDataSourceContext { FolderId = string.Empty }, 
            CancellationToken.None
        );

        // Assert
        foreach (var item in result)
            Console.WriteLine($"Item: {item.DisplayName}, Id: {item.Id}, Type: {(item is Folder ? "Folder" : "File")}");
        Assert.IsNotNull(result);
    }

    [TestMethod]
    public async Task FolderPickerDataHandler_IsSuccess()
    {
        // Arrange
        var folder = new FolderIdentifier { DriveId = "b!vxtzBEWI9Eq4Vkw6hbeSgU7zvg6s6hJLqLwvmwOdy5MXAdvPapxlQKKjxaj1bN1o" };
        var handler = new FolderPickerDataSourceHandler(InvocationContext, folder);

        // Act
        var result = await handler.GetFolderContentAsync(
            new FolderContentDataSourceContext { FolderId = string.Empty }, 
            CancellationToken.None
        );

        // Assert
        foreach (var item in result)
            Console.WriteLine($"Item: {item.DisplayName}, Id: {item.Id}, Type: {(item is Folder ? "Folder" : "File")}");
        Assert.IsNotNull(result);
    }

    [TestMethod]
    public async Task DriveDataSourceHandler_IsSuccess()
    {
        // Arrange
        var handler = new DriveDataSourceHandler(InvocationContext);

        // Act
        var result = await handler.GetDataAsync(new DataSourceContext { SearchString = "" }, CancellationToken.None);

        // Assert
        foreach (var item in result)
            Console.WriteLine($"ID: {item.Value}, Name: {item.DisplayName}");
        Assert.IsNotNull(result);
    }
}
