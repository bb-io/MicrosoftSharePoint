﻿using Apps.MicrosoftSharePoint.Dtos;

namespace Apps.MicrosoftSharePoint.Webhooks.Payload;

public class EventPayload
{
    public SubscriptionDto Subscription { get; set; }
    public string DeltaToken { get; set; }
}