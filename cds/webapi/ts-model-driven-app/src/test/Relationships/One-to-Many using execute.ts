// Demonstrates the following technique:
//  1. Associating two records over a one-to-many relationship
//  2. Disassociating the two records
//  NOTE: This sample uses the execute method with a operationName of 'Associate' and 'Disassociate'

describe("", function() {
  it("One to Many using execute", async function() {
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
          id: contactid.id,
          entityType: "contact"
        };
        relatedEntities = [
          {
            id: accountid.id,
            entityType: "account"
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
      var dissassociateRequest = new class {
        target = {
          id: contactid.id,
          entityType: "contact"
        };
        relatedEntityId = accountid.id;
        relationship = "account_primary_contact";
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
