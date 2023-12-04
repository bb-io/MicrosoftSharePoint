namespace Apps.MicrosoftSharePoint.Models.Dtos;

public class SubscriptionDto
{
    public string Id { get; set; }
    public string Resource { get; set; }
    public string ChangeType { get; set; }
    public DateTime ExpirationDateTime { get; set; }
    public string? NotificationUrl { get; set; }
}

public class SubscriptionWrapper
{
    public IEnumerable<SubscriptionDto> Value { get; set; }
}