// Demonstrates the following technique:
//  Updating a record
//  See: https://docs.microsoft.com/en-us/dynamics365/customer-engagement/developer/webapi/create-entity-web-api

describe("", function() {
  it("Update", async function() {
    this.timeout(90000);
    var assert = chai.assert;

    var record: {
      accountid?: string;
      name?: string;
      address1_city?: string;
    };
    
    record =  {
      name: "Sample Account"
    };

    try {
      // Create Account
      record.accountid = (await (<any>(
        Xrm.WebApi.createRecord("account", record)
      ))).id;

      if (!record.accountid)
      {
        throw new Error("Account ID not returned");
      }
      record.name = "Sample Account (updated)";
      record.address1_city = "Oxford";

      // Update account
      await Xrm.WebApi.updateRecord("account",record.accountid, record);

      // Check the record is updated
      var accountsRead = await Xrm.WebApi.retrieveRecord(
        "account",
        record.accountid,
        "?$select=name,address1_city"
      );
      
      if (!accountsRead || !accountsRead.name || !accountsRead.address1_city) {
        throw new Error("Account not updated");
      }
      assert.equal(accountsRead.name, record.name, "Account updated");
      assert.equal(accountsRead.address1_city, record.address1_city, "Account updated");

    } finally {
      // Delete account
      if (record.accountid) {
        await Xrm.WebApi.deleteRecord("account", record.accountid);
      }
    }
  });
});
