namespace Apps.MicrosoftSharePoint.Models.Responses;

public class ListResponse<T>
{
    public IEnumerable<T> Value { get; set; }
}