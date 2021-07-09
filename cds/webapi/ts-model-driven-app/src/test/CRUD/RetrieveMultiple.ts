// Demonstrates the following technique:
//  Querying for multiple records using odata query
//  See: https://docs.microsoft.com/en-us/dynamics365/customer-engagement/developer/webapi/query-data-web-api

describe("", function() {
  it("RetrieveMultiple", async function() {
    var assert = chai.assert;
    var accountid: {
      entityType: string;
      id: string;
    } = await (<any>Xrm.WebApi.createRecord("account", {
      name: "Sample Account",
      revenue: 20000.01
    }));

    try {
      var results = await Xrm.WebApi.retrieveMultipleRecords(
        "account",
        "?$select=name&$filter=revenue gt 20000 and revenue lt 20001 and name eq 'Sample Account'",
        10
      );

      // Check that there is a single result returned
      if (!results.entities || !results.entities.length)
        throw new Error("No results returned");

      assert.equal(results.entities.length, 1, "Single result returned");
    } finally {
      // Delete account
      if (accountid.id) {
        await Xrm.WebApi.deleteRecord("account", accountid.id);
      }
    }
  });
});
