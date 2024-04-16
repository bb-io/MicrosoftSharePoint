using Newtonsoft.Json.Linq;

namespace Apps.MicrosoftSharePoint.Models.Responses.Pages;

public class PageContentResponse
{
    public string Name { get; set; }
    
    public string Title { get; set; }
    
    public JObject TitleArea { get; set; }
    
    public JObject CanvasLayout { get; set; }
}