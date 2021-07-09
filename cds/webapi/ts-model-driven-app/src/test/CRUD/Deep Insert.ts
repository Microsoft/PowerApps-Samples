// Demonstrates the following technique:
//  Creating an account and 2 related contacts in the same transaction
//  See: https://docs.microsoft.com/en-us/dynamics365/customer-engagement/developer/webapi/associate-disassociate-entities-using-web-api#associate-entities-on-create

describe("", function() {
  it("Deep Insert", async function() {
    this.timeout(90000);
    var assert = chai.assert;
    var context = GetGlobalContext();
    var account: {
      accountid?: string;
      name?: string;
      contact_customer_accounts : {
        firstname: string;
        lastname: string;
      }[];
    };

    account = {
      name: "Sample Account",
      contact_customer_accounts: [
        {
          firstname: "Sample",
          lastname: "contact 1"
        },
        {
          firstname: "Sample",
          lastname: "Contact 2"
        },
      ]
    };

    try {
      // Create Account & Contacts
      account.accountid = (await (<any>(
        Xrm.WebApi.createRecord("account", account)
      ))).id;

      if (!account.accountid)
        throw new Error("Account not created");

      var accountCreated = await Xrm.WebApi.retrieveRecord("account",account.accountid,"?$select=name&$expand=contact_customer_accounts($select=firstname,lastname)");
     
      assert.equal(accountCreated.contact_customer_accounts.length, 2, "Account created with 2 contacts");

    } finally {
      // Delete account
      if (account.accountid) {
        await Xrm.WebApi.deleteRecord("account", account.accountid);
      }
    }
  });
});
