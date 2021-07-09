// This is required due to a bug where entity metadata is not loaded correctly in webresources that open in a new window
// Not needed if the webresource is embdedded inside a form
var windowStatic :any= window;
windowStatic.ENTITY_SET_NAMES = `{
    "account": "accounts",
    "contact": "contacts",
    "opportunity": "opportunities"
}`;

windowStatic.ENTITY_PRIMARY_KEYS = `{
    "account" : "accountid",
    "contact" : "contactid",
    "opportunity" : "opportunityid"
}`;
