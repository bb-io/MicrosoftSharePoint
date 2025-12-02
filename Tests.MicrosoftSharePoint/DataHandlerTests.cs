using Tests.MicrosoftSharePoint.Base;
using Apps.MicrosoftSharePoint.DataSourceHandlers;
using Apps.MicrosoftSharePoint.Models.Requests.Pages;
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
        var handler = new FilePickerDataSourceHandler(InvocationContext);

        // Act
        var result = await handler.GetFolderContentAsync(
            new FolderContentDataSourceContext { FolderId = "root" }, 
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

    [TestMethod]
    public async Task PageDataHandler_IsSuccess()
    {
        // Arrange
        var page = new PageRequest { PageId = "" };
        var handler = new PageDataHandler(InvocationContext, page);

        // Act
        var result = await handler.GetDataAsync(new DataSourceContext { SearchString = "" }, CancellationToken.None);

        // Assert
        foreach (var item in result)
            Console.WriteLine($"ID: {item.Value}, Name: {item.DisplayName}");
        Assert.IsNotNull(result);
    }
}
