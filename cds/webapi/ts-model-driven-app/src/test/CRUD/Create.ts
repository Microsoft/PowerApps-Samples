// Demonstrates the following technique:
//  Creating a record

describe("", function() {
  it("Create", async function() {
    this.timeout(90000);
    var assert = chai.assert;
    var context = GetGlobalContext();

    var account1: {
      accountid?: string;
      name?: string;
      accountcategorycode: Number; //Optionset
      creditlimit: Number; // Money
      creditonhold: Boolean; // Boolean
      numberofemployees: Number; // Integer
      lastonholdtime: Date; // Date
      [index: string]: any; // Allow setting @odata properties
    };

    account1 = {
      name: "Sample Account",
      accountcategorycode: 1, //Preferred_Customer
      creditlimit: 1000,
      creditonhold: true,
      numberofemployees: 10,
      lastonholdtime: new Date(),
      "preferredsystemuserid@odata.bind": `systemusers(${context
        .getUserId()
        .replace("{", "")
        .replace("}", "")})`
    };

    try {
      // Create Account
      account1.accountid = (await (<any>(
        Xrm.WebApi.createRecord("account", account1)
      ))).id;

      if (!account1.accountid) {
        throw new Error("Account not created");
      }

      // Check the account has been created
      var account2 = await Xrm.WebApi.retrieveRecord(
        "account",
        account1.accountid,
        "?$select=name"
      );

      assert.equal(account2.name, "Sample Account", "Account created");
    } finally {
      // Delete account
      if (account1.accountid) {
        await Xrm.WebApi.deleteRecord("account", account1.accountid);
      }
    }
  });
});
