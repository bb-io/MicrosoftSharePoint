using Newtonsoft.Json;

namespace Apps.MicrosoftSharePoint.Models.Dtos.SitePages;

public class NewsPostContentDto : BaseSitePageDto
{
    [JsonProperty("eTag")]
    public string ETag { get; set; }
    public DateTime LastModifiedDateTime { get; set; }
    public string WebUrl { get; set; }
    public string PageLayout { get; set; }
    public string ThumbnailWebUrl { get; set; }
    public bool ShowComments { get; set; }
    public bool ShowRecommendedPages { get; set; }
    public object ContentType { get; set; }
    public object CreatedBy { get; set; }
    public object LastModifiedBy { get; set; }
    public object ParentReference { get; set; }
    public object PublishingState { get; set; }
    public TitleArea TitleArea { get; set; }
    public CanvasLayout CanvasLayout { get; set; }
}

public class TitleArea
{
    public bool EnableGradientEffect { get; set; }
    public string ImageWebUrl { get; set; }
    public string Layout { get; set; }
    public bool ShowAuthor { get; set; }
    public bool ShowPublishedDate { get; set; }
    public bool ShowTextBlockAboveTitle { get; set; }
    public string TextAboveTitle { get; set; }
    public string TextAlignment { get; set; }
    public int ImageSourceType { get; set; }
    public string Title { get; set; }
    public object Authors { get; set; }
    public object AuthorByline { get; set; }
    public string TitlePlaceholder { get; set; }
    public bool IsDecorative { get; set; }
    public ServerProcessedContent ServerProcessedContent { get; set; }
}

public class CanvasLayout
{
    public List<HorizontalSection> HorizontalSections { get; set; }
    public VerticalSection? VerticalSection { get; set; }
}

public class HorizontalSection
{
    public string Layout { get; set; }
    public string Id { get; set; }
    public string Emphasis { get; set; }
    public List<Column> Columns { get; set; }
}

public class Column
{
    public string Id { get; set; }
    public int Width { get; set; }
    public List<Webpart> Webparts { get; set; }
}

public class Webpart
{
    public string Id { get; set; }
    public string? InnerHtml { get; set; }
    public string? WebPartType { get; set; }
    public Data? Data { get; set; }
}

public class VerticalSection
{
    public string Emphasis { get; set; }
    public List<Webpart> Webparts { get; set; }
}

public class ServerProcessedContent
{
    public List<KeyValue> HtmlStrings { get; set; }
    public List<KeyValue> SearchablePlainTexts { get; set; }
    public object Links { get; set; }
    public object ImageSources { get; set; }
    public object? CustomMetadata { get; set; }
    public object? ComponentDependencies { get; set; }
}

public class KeyValue
{
    public string Key { get; set; }
    public string Value { get; set; }
}

public class Data
{
    public List<object> Audiences { get; set; }
    public string DataVersion { get; set; }
    public string Description { get; set; }
    public string Title { get; set; }
    public object Properties { get; set; }
    public ServerProcessedContent? ServerProcessedContent { get; set; }
}