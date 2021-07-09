// Demonstrates the following technique:
//  1. Associating two records over a many to many relationship
//  2. Disassociating the two records
//  See: https://docs.microsoft.com/en-us/dynamics365/customer-engagement/developer/webapi/associate-disassociate-entities-using-web-api

describe("", function() {
  it("Many to Many", async function() {
    this.timeout(90000);
    var assert = chai.assert;

    // Create account
    var accountid: {
      entityType: string;
      id: string;
    } = await (<any>Xrm.WebApi.createRecord("account", {
      name: "Sample Account"
    }));

    // Create lead
    var leadid: {
      entityType: string;
      id: string;
    } = await (<any>Xrm.WebApi.createRecord("lead", {
      lastname: "Sample Lead"
    }));

    try {
      // Associate ManyToMany
      var associate = {
        "@odata.context": WebApiRequest.getOdataContext(),
        "@odata.id": `leads(${leadid.id})`
      };
      var url = `/accounts(${accountid.id})/accountleads_association/$ref`;
      var response = await WebApiRequest.request("POST", url, associate);

      // Check that the association has been made
      var fetch = `<fetch no-lock="true" >
      <entity name="account" >
        <attribute name="name" />
        <filter>
          <condition attribute="accountid" operator="eq" value="${accountid.id}" />
        </filter>
        <link-entity name="accountleads" from="accountid" to="accountid" intersect="true" >
          <attribute name="name" />
        </link-entity>
      </entity>
    </fetch>`;

      var associatedLeads = await Xrm.WebApi.retrieveMultipleRecords(
        "account",
        "?fetchXml=" + fetch
      );

      assert.equal(associatedLeads.entities.length, 1, "Associated records");

      // Disassociate OneToMany
      var url = `/accounts(${accountid.id})/accountleads_association(${
        leadid.id
      })/$ref`;
      var response = await WebApiRequest.request("DELETE", url);
    } finally {
      if (leadid) {
        await Xrm.WebApi.deleteRecord("lead", leadid.id);
      }
      // Delete account
      if (accountid) {
        await Xrm.WebApi.deleteRecord("account", accountid.id);
      }
    }
  });
});
