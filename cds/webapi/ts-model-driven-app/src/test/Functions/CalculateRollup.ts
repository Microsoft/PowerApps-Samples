// Demonstrates the following techniques:
//  Using the CalculateRollup Unbound Function to recalculate a rollup field
//  See: https://docs.microsoft.com/en-us/dynamics365/customer-engagement/developer/webapi/use-web-api-functions

describe("", function() {
  it("CalculateRollup", async function() {
    this.timeout(90000);
    var assert = chai.assert;


    // Get an account
    var response: {
      entities: {
        accountid: String;
        name: String;
      }[];
      nextLink: string;
    } = await Xrm.WebApi.retrieveMultipleRecords(
      "account",
      "?$select=accountid,name&$top=1",
      1
    );

    // Execute request
    var request = new class {
      getMetadata(): any {
        return {
          parameterTypes: {
            FieldName: {
              typeName: "Edm.String",
              structuralProperty: 1
            },
            Target: {
              typeName: "mscrm.crmbaseentity",
              structuralProperty: 5
            }
          },
          operationType: 1,
          operationName: "CalculateRollupField"
        };
      }
      Target = {
        id: response.entities[0].accountid,
        entityType: "account"
      };
      FieldName = "opendeals";
    }();
    await Xrm.WebApi.online.execute(request);
  });
});
