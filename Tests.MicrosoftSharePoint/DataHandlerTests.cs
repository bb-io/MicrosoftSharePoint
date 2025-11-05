using Apps.MicrosoftSharePoint.DataSourceHandlers;
using Blackbird.Applications.Sdk.Common.Dynamic;
using Blackbird.Applications.SDK.Extensions.FileManagement.Models.FileDataSourceItems;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Tests.MicrosoftSharePoint
{
    [TestClass]
    public class DataHandlerTests : TestBase
    {
        [TestMethod]
        public async Task FolderDataHandler_IsSuccess()
        {
            var dataSourceHandler = new FolderDataSourceHandler(InvocationContext);
            var context = new DataSourceContext
            {
                SearchString = ""
            };
            var result = await dataSourceHandler.GetDataAsync(context, CancellationToken.None);

            foreach (var folder in result)
            {
                Console.WriteLine($"Key: {folder.Key}, Value: {folder.Value}");
            }

            Assert.IsNotNull(result);
            Assert.IsTrue(result.Count > 0);
        }

        [TestMethod]
        public async Task FileDataHandler_IsSuccess()
        {
            var dataSourceHandler = new FileDataSourceHandler(InvocationContext);
            var context = new DataSourceContext
            {
                SearchString = "2.srt"
            };
            var result = await dataSourceHandler.GetDataAsync(context, CancellationToken.None);

            foreach (var folder in result)
            {
                Console.WriteLine($"Key: {folder.DisplayName}, Value: {folder.Value}");
            }

            Assert.IsNotNull(result);
        }

        [TestMethod]
        public async Task FilePickerDataHandler_IsSuccess()
        {
            var handler = new FilePickerDataSourceHandler(InvocationContext);
            var result = await handler.GetFolderContentAsync(new FolderContentDataSourceContext
            {
                FolderId = string.Empty
            }, CancellationToken.None);
            var itemList = result.ToList();
            foreach (var item in itemList)
            {
                Console.WriteLine($"Item: {item.DisplayName}, Id: {item.Id}, Type: {(item is Blackbird.Applications.SDK.Extensions.FileManagement.Models.FileDataSourceItems.Folder ? "Folder" : "File")}");
            }
            Assert.IsNotNull(result);
            Assert.IsTrue(itemList.Count > 0, "The folder should contain items.");
        }
    }
}
