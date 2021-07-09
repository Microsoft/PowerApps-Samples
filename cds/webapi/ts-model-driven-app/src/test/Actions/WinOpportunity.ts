// Demonstrates the following techniques:
//  1. Creating an opportunity
//  2. Winning an opportunity using the execute method
// See: https://docs.microsoft.com/en-us/dynamics365/customer-engagement/developer/webapi/use-web-api-functions

describe("", function() {

  it("Win Opportunity", async function() {
    this.timeout(90000);
    var assert = chai.assert;
    
    var account1: any = {
      name: "Sample Account"
    };

    var createAccountResponse: {
      entityType: String;
      id: String;
    } = <any>await Xrm.WebApi.createRecord("account", account1);

    account1.accountid = createAccountResponse.id;

    var opportunity1: any = {
      name: "Sample Opportunity",
      estimatedvalue: 1000,
      estimatedclosedate: "2019-02-10",
      "parentaccountid@odata.bind": `accounts(${account1.accountid})`
    };

    var createOpportunityResponse: {
      entityType: String;
      id: String;
    } = <any>await Xrm.WebApi.createRecord("opportunity", opportunity1);

    opportunity1.opportunityid = createOpportunityResponse.id;

    // Execute request
    var winOpportunityRequest = new class {
      OpportunityClose = {
        description: "Sample Opportunity Close",
        subject: "Sample",
        "@odata.type": "Microsoft.Dynamics.CRM.opportunityclose",
        "opportunityid@odata.bind": `opportunities(${
          opportunity1.opportunityid
        })`
      };
      Status = 3;

      getMetadata(): any {
        return {
          parameterTypes: {
            OpportunityClose: {
              typeName: "mscrm.opportunityclose",
              structuralProperty: 5
            },
            Status: {
              typeName: "Edm.Int32",
              structuralProperty: 1
            }
          },
          operationType: 0,
          operationName: "WinOpportunity"
        };
      }
    }();
    var rawResponse = <any>(
      await Xrm.WebApi.online.execute(winOpportunityRequest)
    );
    var response = await rawResponse.text();

    // Delete account
    if (response.id) {
      await Xrm.WebApi.deleteRecord("account", response.id);
    }
  });
});
