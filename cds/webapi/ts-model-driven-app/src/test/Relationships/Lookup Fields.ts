// Demonstrates the following techniques:
//  1. Creating a record with a lookup field
//  2. Updating a lookup field on an existing record
//  3. Setting a lookup field to null on an existing record
//  4. Querying Lookups

/// <reference path="../WebApiRequest.ts" />
describe("", function() {
  it("Lookup Fields", async function() {
    this.timeout(90000);

    // Create a contact
    var contact1id: {
      entityType: string;
      id: string;
    } = await (<any>Xrm.WebApi.createRecord("contact", {
      lastname: "Sample Contact 1"
    }));

    var contact2id: {
      entityType: string;
      id: string;
    } = await (<any>Xrm.WebApi.createRecord("contact", {
      lastname: "Sample Contact 2"
    }));

    // Create account
    var accountid: {
      entityType: string;
      id: string;
    } = await (<any>Xrm.WebApi.createRecord("account", {
      name: "Sample Account",
      "primarycontactid@odata.bind": `/contacts(${contact1id.id})`
    }));

    try {
      // Update the primary contact to a different contact
      await Xrm.WebApi.updateRecord("account",accountid.id, {
        "primarycontactid@odata.bind": `/contacts(${contact2id.id})`

      });

      // Disassociate Contact to Account as Primary Contact
      // Note:  It is not possible to update a lookup field to be null
      //        Each field beign nulled must have a separate DELETE request 
      var url = `/accounts(${accountid.id})/primarycontactid/$ref`;
      var response = await WebApiRequest.request("DELETE",url);

    } finally {
      // Delete account
      if (accountid) {
        await Xrm.WebApi.deleteRecord("account", accountid.id);
      }

      // Delete contact 1
      if (contact1id) {
        await Xrm.WebApi.deleteRecord("contact", contact1id.id);
      }

      // Delete contact 2
      if (contact2id) {
        await Xrm.WebApi.deleteRecord("contact", contact2id.id);
      }
    }
  });
});
