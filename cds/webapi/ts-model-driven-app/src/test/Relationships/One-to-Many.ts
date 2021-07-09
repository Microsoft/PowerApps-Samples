/// <reference path="../WebApiRequest.ts" />
// Demonstrates the following technique:
//  1. Associating two records over a one-to-many relationship
//  2. Disassociating the two records
//  See: https://docs.microsoft.com/en-us/dynamics365/customer-engagement/developer/webapi/associate-disassociate-entities-using-web-api

describe("", function() {
  it("One to Many", async function() {
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
        "@odata.id": `accounts(${accountid.id})`
      };
      var url = `/contacts(${contactid.id})/account_primary_contact/$ref`;
      var response = await WebApiRequest.request("POST", url, associate);

      // Disassociate Contact to Account as Primary Contact
      var url = `/contacts(${contactid.id})/account_primary_contact(${accountid.id})/$ref`;
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
