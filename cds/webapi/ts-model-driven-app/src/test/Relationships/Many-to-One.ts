/// <reference path="../WebApiRequest.ts" />
// Demonstrates the following technique:
//  1. Associating two records over a many-to-one relationship. This is an alternative to simply using a lookup field in a create/update
//  2. Disassociating the two records
//  See: https://docs.microsoft.com/en-us/dynamics365/customer-engagement/developer/webapi/associate-disassociate-entities-using-web-api

describe("", function() {
  it("Many to One", async function() {
    this.timeout(90000);
    // Create a contact
    var contactid: {
      entityType: string;
      id: string;
    } = await (<any>Xrm.WebApi.createRecord("contact", {
      lastname: "Sample Contact"
    }));

    // Create account
    var accountid: {
      entityType: string;
      id: string;
    } = await (<any>Xrm.WebApi.createRecord("account", {
      name: "Sample Account"
    }));

    try {
      // Associate Contact to Account as Primary Contact
      var associate = {
        "@odata.context": WebApiRequest.getOdataContext(),
        "@odata.id": `contacts(${contactid.id})`
      };
      var url = `/accounts(${accountid.id})/primarycontactid/$ref`;
      var response = await WebApiRequest.request("PUT", url, associate);

      // Disassociate Contact to Account as Primary Contact
      // Note: When disassociating Many to One - the lookup attribute name is used with only the left hand side id
      var url = `/accounts(${accountid.id})/primarycontactid/$ref`;
      var response = await WebApiRequest.request("DELETE",url);

    } finally {
      // Delete account
      if (accountid) {
        await Xrm.WebApi.deleteRecord("account", accountid.id);
      }

      // Delete contact
      if (contactid) {
        await Xrm.WebApi.deleteRecord("contact", contactid.id);
      }
    }
  });
});
