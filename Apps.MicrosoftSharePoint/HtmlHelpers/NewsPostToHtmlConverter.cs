using System.Text;
using Apps.MicrosoftSharePoint.Models.Dtos.SitePages;

namespace Apps.MicrosoftSharePoint.HtmlHelpers;

public static class NewsPostToHtmlConverter
{
    const string DataIdAttribute = "data-id";
    
    public static string ConvertToHtml(this NewsPostContentDto post)
    {
        var html = new StringBuilder();

        html.Append($"<div> {DataIdAttribute}={post.Id}");

        #region Title area

        html.Append($"<div {DataIdAttribute}=\"{nameof(post.TitleArea)}\">");
        html.Append($"<div {DataIdAttribute}=\"{nameof(post.TitleArea.Title)}\">{post.TitleArea.Title}</div>");
        html.Append($"<div {DataIdAttribute}=\"{nameof(post.TitleArea.TextAboveTitle)}\">{post.TitleArea.TextAboveTitle}</div>");
        
        var titleServerProcessedContentHtml = ConvertServerProcessedContent(post.TitleArea.ServerProcessedContent);
        html.Append(titleServerProcessedContentHtml);

        html.Append("</div>");

        #endregion

        #region Canvas layout

        html.Append($"<div {DataIdAttribute}=\"{nameof(post.CanvasLayout)}\">");

        #region Horizontal sections

        html.Append($"<div {DataIdAttribute}=\"{nameof(post.CanvasLayout.HorizontalSections)}\">");

        foreach (var section in post.CanvasLayout.HorizontalSections)
        {
            html.Append($"<div {DataIdAttribute}=\"{section.Id}\">");

            foreach (var column in section.Columns)
            {
                html.Append($"<div {DataIdAttribute}=\"{column.Id}\">");

                foreach (var webpart in column.Webparts)
                {
                    html.Append($"<div {DataIdAttribute}=\"{webpart.Id}\">");
                    
                    if (webpart.InnerHtml != null)
                        html.Append($"<div {DataIdAttribute}=\"{nameof(webpart.InnerHtml)}\">{webpart.InnerHtml}</div>");

                    if (webpart.Data?.ServerProcessedContent != null)
                    {
                        var webpartServerProcessedContentHtml = ConvertServerProcessedContent(webpart.Data.ServerProcessedContent);
                        html.Append(webpartServerProcessedContentHtml);
                    }

                    html.Append("</div>");
                }
                
                html.Append("</div>");
            }
            
            html.Append("</div>");
        }
        
        html.Append("</div>");

        #endregion

        #region Vertical section

        if (post.CanvasLayout.VerticalSection != null)
        {
            html.Append($"<div {DataIdAttribute}=\"{nameof(post.CanvasLayout.VerticalSection)}\">");

            foreach (var webpart in post.CanvasLayout.VerticalSection.Webparts)
            {
                html.Append($"<div {DataIdAttribute}=\"{webpart.Id}\">");
                    
                if (webpart.InnerHtml != null)
                    html.Append($"<div {DataIdAttribute}=\"{nameof(webpart.InnerHtml)}\">{webpart.InnerHtml}</div>");

                if (webpart.Data?.ServerProcessedContent != null)
                {
                    var webpartServerProcessedContentHtml = ConvertServerProcessedContent(webpart.Data.ServerProcessedContent);
                    html.Append(webpartServerProcessedContentHtml);
                }

                html.Append("</div>");
            }
        
            html.Append("</div>");
        }

        #endregion
        
        html.Append("</div>");

        #endregion
        
        html.Append("</div>");

        return html.ToString();
    }

    private static string ConvertServerProcessedContent(ServerProcessedContent content)
    {
        var html = new StringBuilder();
        
        html.Append($"<div {DataIdAttribute}=\"{nameof(ServerProcessedContent)}\">");
                        
        html.Append($"<div {DataIdAttribute}=\"{nameof(content.HtmlStrings)}\">");
        foreach (var htmlString in content.HtmlStrings)
        {
            html.Append($"<div {DataIdAttribute}=\"{htmlString.Key}\">{htmlString.Value}</div>");
        }
        html.Append("</div>");

        html.Append($"<div {DataIdAttribute}=\"{nameof(content.SearchablePlainTexts)}\">");
        foreach (var text in content.SearchablePlainTexts)
        {
            html.Append($"<div {DataIdAttribute}=\"{text.Key}\">{text.Value}</div>");
        }
        html.Append("</div>");
                        
        html.Append("</div>");

        return html.ToString();
    }
}