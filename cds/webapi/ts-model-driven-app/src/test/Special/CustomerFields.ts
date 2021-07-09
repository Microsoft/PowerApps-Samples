// Demonstrates the following techniques:
//  Creating records with customer fields that reference either an account or a contact

/// <reference path="../WebApiRequest.ts" />
describe("", function() {
  it("Customer Fields", async function() {
    this.timeout(90000);

    console.log("Creating account");
    var assert = chai.assert;

    // Create account
    const account1 = {
      name: `Sample Account ${new Date().toUTCString()}`
    };
    var account1id = (<any>await Xrm.WebApi.createRecord("account", account1))
      .id;

    if (!account1id) throw new Error("accountid not defined");

    // Create contact
    const contact1 = {
      lastname: `Sample Contact ${new Date().toUTCString()}`
    };
    var contact1id = (<any>await Xrm.WebApi.createRecord("contact", contact1))
      .id;

    if (!contact1id) throw new Error("contact1id not defined");

    try {
      // Create opportunity for the created account
      const opportunity1: any = {
        name: `Sample Opportunity ${new Date().toUTCString()}`,
        estimatedvalue: 1000,
        estimatedclosedate: new Date(Date.now()).toISOString().substr(0, 10), // DateOnly
        "customerid_account@odata.bind": `accounts(${account1id})`
      };

      var opportunity1id = (<any>(
        await Xrm.WebApi.createRecord("opportunity", opportunity1)
      )).id;

      if (!opportunity1id) throw new Error("Opportunty ID not defined");

      // Retrieve the opportunity
      var opportunity2 = await Xrm.WebApi.retrieveRecord(
        "opportunity",
        opportunity1id,
        "?$select=name,_customerid_value"
      );

      if (!opportunity2 || !opportunity2._customerid_value)
        throw new Error("Opportunity2 Customerid not returned");

      // Check that the customerid field is populated
      assert.isNotEmpty(
        opportunity2._customerid_value,
        "Customer field not empty"
      );

      assert.equal(
        opportunity2._customerid_value,
        account1id,
        "Customer id equal to accountid"
      );

      assert.equal(
        opportunity2[
          "_customerid_value@Microsoft.Dynamics.CRM.lookuplogicalname"
        ],
        "account",
        "Logical name set"
      );

      // Update the customer field to reference the contact
      delete opportunity1["customerid_account@odata.bind"];
      opportunity1["customerid_contact@odata.bind"] = `contacts(${contact1id})`;

      await Xrm.WebApi.updateRecord(
        "opportunity",
        opportunity1id,
        opportunity1
      );

      // Retrieve the opportunity
      var opportunity3 = await Xrm.WebApi.retrieveRecord(
        "opportunity",
        opportunity1id,
        "?$select=name,_customerid_value"
      );

      // Check that the customerid field is populated
      assert.isNotEmpty(
        opportunity3._customerid_value,
        "Customer field not empty"
      );
      assert.equal(
        opportunity3._customerid_value,
        contact1id,
        "Customer id equal to contact"
      );

      assert.equal(
        opportunity3[
          "_customerid_value@Microsoft.Dynamics.CRM.lookuplogicalname"
        ],
        "contact",
        "Logical name set"
      );
    } finally {
      // Delete the opportunity and account - opportunity is a cascade delete
      await Xrm.WebApi.deleteRecord("contact", contact1id);
      await Xrm.WebApi.deleteRecord("account", account1id);
    }
  });
});
