// Demonstrates the following technique:
//  Creating a record and then retrieving it unchanged with the same etag
//  This is a technique employed by the client side api to avoid uneccesary data being transferred over the network
//  This sample shows how varying the options will invalidate the client side cache
//  See: https://docs.microsoft.com/en-us/dynamics365/customer-engagement/developer/webapi/create-entity-web-api

describe("", function() {
  it("etags with $expand", async function() {
    this.timeout(90000);

    var assert = chai.assert;

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
      name: "Sample Account",
      "primarycontactid@odata.bind": `/contacts(${contactid.id})`
    }));

    try {
      // Read the Account
      var query1 = await Xrm.WebApi.retrieveRecord(
        "account",
        accountid.id,
        "?$select=name&$expand=primarycontactid"
      );
      var etag1 = query1["@odata.etag"];

      // Read the Account again (This will return the same record as above since the server will return 304 Not Modified )
      var query2 = await Xrm.WebApi.retrieveRecord(
        "account",
        accountid.id,
        "?$select=name&$expand=primarycontactid"
      );
      var etag2 = query2["@odata.etag"];

      assert.equal(etag1, etag2, "Record not modified");
      assert.equal(
        "Sample Contact",
        query2.primarycontactid.lastname,
        "Related contact returned"
      );

      // Update the contact name
      await (<any>Xrm.WebApi.updateRecord("contact", contactid.id, {
        lastname: "Sample Contact (edited)"
      }));

      // Read the Account again. Since only the related expanded record was updated, the retrieve will not return the correct value.
      var query3 = await Xrm.WebApi.retrieveRecord(
        "account",
        accountid.id,
        "?$select=name&$expand=primarycontactid"
      );
      var etag3 = query3["@odata.etag"];

      assert.equal(etag1, etag3, "Record not modified");
      assert.equal(
        "Sample Contact",
        query3.primarycontactid.lastname,
        "Unchanged contact"
      );

      // Workaround: Changing the $select query will result in a client side cache miss and the new value will be returned
      var query4 = await Xrm.WebApi.retrieveRecord(
        "account",
        accountid.id,
        "?$select=name&$expand=primarycontactid($select=lastname)"
      );
      var etag4 = query4["@odata.etag"];
   
      assert.equal(etag1, etag4, "Record not modified");
      assert.equal(
        "Sample Contact (edited)",
        query4.primarycontactid.lastname,
        "Contact changed"
      );
    } finally {
      // Delete account & contact
      if (accountid.id) {
        await Xrm.WebApi.deleteRecord("account", accountid.id);
      }
      if (contactid.id) {
        await Xrm.WebApi.deleteRecord("contact", contactid.id);
      }
    }
  });
});
