using Apps.MicrosoftSharePoint.DataSourceHandlers;
using Blackbird.Applications.Sdk.Common.Dynamic;
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
    }
}
