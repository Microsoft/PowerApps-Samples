// Demonstrates the following technique:
//  Querying for a record by id using retrieveRecord with a $select clause

describe("", function() {
  it("Read", async function() {
    this.timeout(90000);
    var assert = chai.assert;

    var record: {
      accountid?: string;
      name?: string;
    };

    record =  {
      name: "Sample Account"
    };

    // Create Account
    record.accountid = (await (<any>(
      Xrm.WebApi.createRecord("account", record)
    ))).id;

    if (!record.accountid)
      throw new Error("Account not created");

    try {
      var accountsRead = await Xrm.WebApi.retrieveRecord(
        "account",
        record.accountid,
        "?$select=name,primarycontactid"
      );

      if (!accountsRead || !accountsRead.name) {
        throw new Error("Account not created");
      }
      assert.equal(accountsRead.name, record.name, "Account created");

    } finally {
      // Delete account
      if (record.accountid) {
        await Xrm.WebApi.deleteRecord("account", record.accountid);
      }
    }
  });
});
