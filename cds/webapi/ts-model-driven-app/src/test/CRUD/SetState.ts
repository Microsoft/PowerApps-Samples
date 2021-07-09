// Demonstrates the following technique:
//  Updating a record's State and Status Resaon
//  See: https://docs.microsoft.com/en-us/dynamics365/customer-engagement/developer/webapi/create-entity-web-api
const enum opportunity_statuscode {
  In_Progress = 1,
  On_Hold = 2,
  Won = 3,
  Canceled = 4,
  OutSold = 5
}
const enum account_statecode {
  Active = 0,
  Inactive = 1
}
describe("", function() {
  it("SetState", async function() {
    this.timeout(90000);
    var assert = chai.assert;

    var account: {
      accountid?: string;
      name?: string;
      address1_city?: string;
      statecode?: account_statecode;
    };

    account = {
      name: "Sample Account"
    };

    try {
      // Create Account
      account.accountid = (await (<any>(
        Xrm.WebApi.createRecord("account", account)
      ))).id;

      if (!account.accountid) {
        throw new Error("Account ID not returned");
      }

      // Create Opportunity for the account
      var opportunity: any = {
        name: "Sample Opportunity",
        "parentaccountid@odata.bind": `accounts(${account.accountid})`
      };

      var createOpportunityResponse: {
        entityType: String;
        id: String;
      } = <any>await Xrm.WebApi.createRecord("opportunity", opportunity);

      opportunity.opportunityid = createOpportunityResponse.id;

      // Change Opportunity Status Reason to In Progress
      opportunity.statuscode = opportunity_statuscode.In_Progress;
      await Xrm.WebApi.updateRecord(
        "opportunity",
        opportunity.opportunityid,
        opportunity
      );

      // Check the opportunity is updated
      var opportunityRead = await Xrm.WebApi.retrieveRecord(
        "opportunity",
        opportunity.opportunityid,
        "?$select=statuscode"
      );

      if (!opportunityRead || !opportunityRead.statuscode) {
        throw new Error("Opportunity not updated");
      }
      assert.equal(
        opportunityRead.statuscode,
        opportunity_statuscode.In_Progress,
        "Opportunity In Progress"
      );

      // Update account state to In Active
      account.statecode = account_statecode.Inactive;
      await Xrm.WebApi.updateRecord("account", account.accountid, account);

      // Check the account is updated
      var accountRead = await Xrm.WebApi.retrieveRecord(
        "account",
        account.accountid,
        "?$select=statecode"
      );

      if (!accountRead || accountRead.statecode == undefined) {
        throw new Error("Account not updated");
      }
      assert.equal(
        accountRead.statecode,
        account_statecode.Inactive,
        "Account Inactive"
      );
    } finally {
      // Delete account
      if (account.accountid) {
        await Xrm.WebApi.deleteRecord("account", account.accountid);
      }
    }
  });
});
