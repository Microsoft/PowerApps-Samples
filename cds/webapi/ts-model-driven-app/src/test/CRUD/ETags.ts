// Demonstrates the following technique:
//  Creating a record and then retrieving it unchanged with the same etag
//  This is a technique employed by the client side api to avoid uneccesary data being transferred over the network
//  See: https://docs.microsoft.com/en-us/dynamics365/customer-engagement/developer/webapi/create-entity-web-api

describe("", function() {
  it("etags", async function() {
    this.timeout(90000);

    var assert = chai.assert;
    
    var account: {
      accountid?: string;
      name?: string;
      creditlimit: Number; // Money
      [index: string]: any; // Allow setting @odata properties
    };

    account = {
      name: "Sample Account",
      creditlimit: 1000,
    };

    try {
      // Create Account
      account.accountid = (await (<any>(
        Xrm.WebApi.createRecord("account", account)
      ))).id;

      if (!account.accountid)
        throw new Error("Account not created");

      // Read the Account
      var query1 = await Xrm.WebApi.retrieveRecord("account",account.accountid,"?$select=name");
      var etag1 = query1["@odata.etag"];

      // Read the Account again (This will return the same record as above since the server will return 304 Not Modified )
      var query2 = await Xrm.WebApi.retrieveRecord("account",account.accountid,"?$select=name");
      var etag2 = query2["@odata.etag"];

      assert.equal(etag1,etag2,"Record not modified");

      // Update the value
      account.name = "Sample Account (updated)";
      await Xrm.WebApi.updateRecord("account",account.accountid, account);

      // Read the Account again. Since the record is updated on the server it will have a different etag
      var query3 = await Xrm.WebApi.retrieveRecord("account",account.accountid,"?$select=name");
      var etag3 = query3["@odata.etag"];

      assert.notEqual(etag1,etag3,"Record modified");


    } finally {
      // Delete account
      if (account.accountid) {
        await Xrm.WebApi.deleteRecord("account", account.accountid);
      }
    }
  });
});
