using System.Globalization;
using System.Text;
using System.Web;
using Apps.MicrosoftSharePoint.Models.Responses.Pages;
using HtmlAgilityPack;
using Newtonsoft.Json.Linq;

namespace Apps.MicrosoftSharePoint.HtmlConversion;

public static class SharePointHtmlConverter
{
    private static readonly string[] TranslatableProperties = ["innerHtml", "label", "searchablePlainTexts"];

    public static byte[] ToHtml(PageContentResponse pageContent)
    {
        var (doc, body) = PrepareEmptyHtmlDocument(pageContent.TitleArea?["title"]!.ToString() ?? pageContent.Title);

        body.SetAttributeValue(ConversionConstants.OriginalAttr, pageContent.CanvasLayout.ToString());

        pageContent.CanvasLayout.Descendants()
            .Where(x => x is JProperty jProperty && TranslatableProperties.Contains(jProperty.Name))
            .Cast<JProperty>()
            .ToList()
            .ForEach(x => AppendChild(body, doc, x));

        return Encoding.UTF8.GetBytes(doc.DocumentNode.OuterHtml);
    }

    private static void AppendChild(HtmlNode body, HtmlDocument doc, JProperty prop)
    {
        switch (prop.Name)
        {
            case "searchablePlainTexts":
                var texts = prop.Value as JArray;
                texts!.ToList().ForEach(x => AppendTextChild(body, doc, (x["value"] as JValue)!));
                break;
            default:
                AppendTextChild(body, doc, (prop.Value as JValue)!);
                break;
        }
    }

    private static void AppendTextChild(HtmlNode body, HtmlDocument doc, JValue value)
    {
        var node = doc.CreateElement(HtmlConstants.Div);

        node.SetAttributeValue(ConversionConstants.PathAttr, value.Path);
        node.InnerHtml = value.ToString(CultureInfo.InvariantCulture);

        body.AppendChild(node);
    }

    private static (HtmlDocument document, HtmlNode bodyNode) PrepareEmptyHtmlDocument(string title)
    {
        var htmlDoc = new HtmlDocument();
        var htmlNode = htmlDoc.CreateElement(HtmlConstants.Html);
        htmlDoc.DocumentNode.AppendChild(htmlNode);

        var headNode = htmlDoc.CreateElement(HtmlConstants.Head);

        var titleNode = htmlDoc.CreateElement(HtmlConstants.Title);
        titleNode.InnerHtml = title;

        headNode.AppendChild(titleNode);
        htmlNode.AppendChild(headNode);

        var bodyNode = htmlDoc.CreateElement(HtmlConstants.Body);
        htmlNode.AppendChild(bodyNode);

        return (htmlDoc, bodyNode);
    }

    public static (string title, JObject content) ToJson(Stream fileStream)
    {
        var doc = new HtmlDocument();
        doc.Load(fileStream);

        var title = HttpUtility.HtmlDecode(doc.DocumentNode.Descendants().First(x => x.Name == HtmlConstants.Title)
            .InnerText);
        var sourceContent = doc.DocumentNode.Descendants()
            .First(x => x.Attributes[ConversionConstants.OriginalAttr]?.Value != null)
            .Attributes[ConversionConstants.OriginalAttr].Value;

        var sourceJObject = JObject.Parse(HttpUtility.HtmlDecode(sourceContent));

        doc.DocumentNode
            .Descendants()
            .Where(x => x.Attributes[ConversionConstants.PathAttr]?.Value != null)
            .ToList()
            .ForEach(x =>
            {
                var path = x.Attributes[ConversionConstants.PathAttr].Value!;
                var propertyValue = sourceJObject.SelectToken(path);

                (propertyValue as JValue)!.Value = HttpUtility.HtmlDecode(x.InnerHtml.Trim());
            });

        return (title, sourceJObject);
    }
}