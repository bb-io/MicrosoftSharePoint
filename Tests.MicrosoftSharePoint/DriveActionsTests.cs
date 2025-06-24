using Apps.MicrosoftSharePoint.Actions;
using Apps.MicrosoftSharePoint.Models.Identifiers;
using Blackbird.Applications.Sdk.Common.Dynamic;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Tests.MicrosoftSharePoint
{
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
        public async Task GetFolder_IsSuccess()
        {
            var action = new DriveActions(InvocationContext, FileManager);

            var result = await action.GetFolderMetadataById(
                new ParentFolderIdentifier { ParentFolderId = "01C7WXPSHVF2MLHRDQM5GJNQXIMR5A3QGW" },
                "Backup");
            Console.WriteLine($"Key: {result.Id}, Value: {result.Name}");

            Assert.IsNotNull(result);
        }
    }
}
