// Demonstrates the following technique:
//  Deleting a record

describe("", function() {
  it("Delete", async function() {
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

    // Delete account
    if (record.accountid) {
      await Xrm.WebApi.deleteRecord("account", record.accountid);
    }

    // Check the account has been deleted
    var fetch = `<fetch no-lock="true" >
       <entity name="account" >
         <filter>
           <condition attribute="accountid" operator="eq" value="${
             record.accountid
           }" />
         </filter>
       </entity>
     </fetch>`;

    var accounts = await Xrm.WebApi.retrieveMultipleRecords(
      "account",
      "?fetchXml=" + fetch
    );

    assert.equal(accounts.entities.length, 0, "Account deleted");
  });
});
