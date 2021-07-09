// Demonstrates the following technique:
//  1. Associating two records over a many-to-many relationship
//  2. Disassociating the two records
//  NOTE: This sample uses the execute method with a operationName of 'Associate' and 'Disassociate'

describe("", function() {
  it("Many to One using execute", async function() {
    this.timeout(90000);
    // Create a contact
    var contactid: {
      entityType: string;
      id: string;
    } = await (<any>Xrm.WebApi.createRecord("contact", {
      lastname: "Sample Contact"
    }));

    // Create account
    var accountid: {
      entityType: string;
      id: string;
    } = await (<any>Xrm.WebApi.createRecord("account", {
      name: "Sample Account"
    }));

    try {
      // Associate Contact to Account as Primary Contact
      var associateRequest = new class {
        target = {
          id: accountid.id,
          entityType: "account"
        };
        relatedEntities = [
          {
            id: contactid.id,
            entityType: "contact"
          }
        ];
        relationship = "account_primary_contact";
        getMetadata(): any {
          return {
            parameterTypes: {},
            operationType: 2,
            operationName: "Associate"
          };
        }
      }();

      var response = await Xrm.WebApi.online.execute(associateRequest);

      // Disassociate Contact to Account as Primary Contact
      // Note: When disassociating Many to One - the lookup attribute name is used with only the left hand side id
      var dissassociateRequest = new class {
        target = {
          id: accountid.id,
          entityType: "account"
        };
        relationship = "primarycontactid";
        getMetadata(): any {
          return {
            parameterTypes: {},
            operationType: 2,
            operationName: "Disassociate"
          };
        }
      }();

      var dissassociateResponse = await Xrm.WebApi.online.execute(
        dissassociateRequest
      );

    } finally {

      // Delete account
      if (accountid) {
        await Xrm.WebApi.deleteRecord("account", accountid.id);
      }

      // Delete contact
      if (contactid) {
        await Xrm.WebApi.deleteRecord("contact", contactid.id);
      }
    }
  });
});
