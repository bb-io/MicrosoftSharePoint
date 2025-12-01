using Apps.MicrosoftSharePoint.DataSourceHandlers;
using Blackbird.Applications.SDK.Extensions.FileManagement.Models.FileDataSourceItems;

namespace Tests.MicrosoftSharePoint;

[TestClass]
public class DataHandlerTests : TestBase
{
    [TestMethod]
    public async Task FilePickerDataHandler_IsSuccess()
    {
        // Arrange
        var handler = new FilePickerDataSourceHandler(InvocationContext);

        // Act
        var result = await handler.GetFolderContentAsync(
            new FolderContentDataSourceContext { FolderId = "b!V1tgT5LcyEiu-qCrLWi_sYE68iUDk7hCsFdpT5k21zfKEJWmu4ZAR7Zk_O7bubtF#017O7UAG7AGGYZEMSZFNDL6GYZFR6R2JGX" }, 
            CancellationToken.None
        );

        // Assert
        foreach (var item in result)
            Console.WriteLine($"Name: {item.DisplayName}, Id: {item.Id}, Type: {(item is Folder ? "Folder" : "File")}");
        Assert.IsNotNull(result);
    }

    [TestMethod]
    public async Task FolderPickerDataHandler_IsSuccess()
    {
        // Arrange
        var handler = new FolderPickerDataSourceHandler(InvocationContext);

        // Act
        var result = await handler.GetFolderContentAsync(
            new FolderContentDataSourceContext { FolderId = string.Empty }, 
            CancellationToken.None
        );

        // Assert
        foreach (var item in result)
            Console.WriteLine($"Name: {item.DisplayName}, Id: {item.Id}, Type: {(item is Folder ? "Folder" : "File")}");
        Assert.IsNotNull(result);
    }
}
