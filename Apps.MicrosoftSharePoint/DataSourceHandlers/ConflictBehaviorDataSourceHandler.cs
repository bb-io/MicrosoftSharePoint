using Blackbird.Applications.Sdk.Common.Dynamic;
using Blackbird.Applications.Sdk.Common.Dictionaries;

namespace Apps.MicrosoftSharePoint.DataSourceHandlers;

public class ConflictBehaviorDataSourceHandler : IStaticDataSourceItemHandler
{
    public IEnumerable<DataSourceItem> GetData()
    {
        return new List<DataSourceItem>
        {
            new DataSourceItem("fail", "Fail uploading"),
            new DataSourceItem("replace", "Replace file"),
            new DataSourceItem("rename", "Rename file"),
        };
    }
}