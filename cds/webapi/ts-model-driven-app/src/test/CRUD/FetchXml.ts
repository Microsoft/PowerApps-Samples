// Demonstrates the following technique:
//  Querying records using fetchxml

describe("", function() {
  it("Query with FetchXml", async function() {
    this.timeout(90000);

    var assert = chai.assert;
    // Check the account has been created
    var fetch = `<fetch no-lock="true" >
      <entity name="account">
        <attribute name="name"/>
      </entity>
    </fetch>`;

    var accounts = await Xrm.WebApi.retrieveMultipleRecords(
      "account",
      "?fetchXml=" + fetch
    );

    assert.isNotNull(accounts.entities, "Account query returns results");
  });
});
