// Demonstrates the following techniques:
//  Using the CalculateTotalTimeIncident Bound Function to recalculate a rollup field
//  See: https://docs.microsoft.com/en-us/dynamics365/customer-engagement/developer/webapi/use-web-api-functions

describe("", function() {
  it("CalculateTotalTimeIncident", async function() {
    this.timeout(90000);
    var assert = chai.assert;

    // Get an incident
    var response: {
      entities: {
        incidentid: String;
      }[];
      nextLink: string;
    } = await Xrm.WebApi.retrieveMultipleRecords(
      "incident",
      "?$select=incidentid",
      1
    );

    // Execute CalculateTotalTimeIncident request
    // This is a bound function which we pass the entity parameter as the target incident
    var request = new class {
      entity = {
        id: response.entities[0].incidentid,
        entityType: "incident"
      };

      getMetadata(): any {
        return {
          boundParameter: "entity",
          parameterTypes: {
            entity: {
              typeName: "mscrm.incident",
              structuralProperty: 5
            }
          },
          operationType: 1,
          operationName: "CalculateTotalTimeIncident"
        };
      }
    }();

    // The json function is a promise which returns the response object.
    var rawResponse = await (<any>await Xrm.WebApi.online.execute(request)).json();;
   
    assert.isNumber(rawResponse.TotalTime,"Total Time returned");
  });
});
