﻿using Newtonsoft.Json;
using Apps.MicrosoftSharePoint.Actions;
using Apps.MicrosoftSharePoint.Models.Requests;
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
            new ParentFolderIdentifier { ParentFolderId = "01C7WXPSHVF2MLHRDQM5GJNQXIMR5A3QGW" },
            "Backup");
            Console.WriteLine($"Key: {result.Id}, Value: {result.Name}");

        Assert.IsNotNull(result);
    }

    [TestMethod]
    public async Task GetFileMetadata_IsSuccess()
    {
        // Arrange
        var action = new DriveActions(InvocationContext, FileManager);
        var fileId = "01Q6TCMYOLVEFH7YXSJBGZKI2JGUATR5AH";

        // Act
        var result = await action.GetFileMetadataById(new FileIdentifier { FileId = fileId });

        // Assert
        Console.WriteLine(JsonConvert.SerializeObject(result, Formatting.Indented));
        Assert.IsNotNull(result);
    }

    [TestMethod]
    public async Task ListFilesInFolderById_IsSuccess()
    {
        // Arrange
        var action = new DriveActions(InvocationContext, FileManager);
        var folderId = "01Q6TCMYPJNODF2KHBNRBYIMJKJZHHSLOR";

        // Act
        var result = await action.ListFilesInFolderById(
            new FolderIdentifier { FolderId = folderId },
            new filterExtensions { Extensions = [".doc", ".docx"] }
        );

        // Assert
        Console.WriteLine(JsonConvert.SerializeObject(result, Formatting.Indented));
        Assert.IsNotNull(result);
    }
}
