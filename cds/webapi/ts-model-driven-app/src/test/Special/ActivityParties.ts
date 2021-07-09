// Demonstrates the following techniques:
//  1. Creating activities with activity parties
//  2. Updating activity parties - this is a special case
//  See: https://docs.microsoft.com/en-us/dynamics365/customer-engagement/developer/webapi/associate-disassociate-entities-using-web-api#associate-entities-on-update-using-collection-valued-navigation-property

/// <reference path="../WebApiRequest.ts" />
describe("", function() {
  it("Activity Parties Letter", async function() {
    this.timeout(900000);
    var assert = chai.assert;

    // Create contact 1
    var contact1 = {
      lastname: `Test Contact 1${new Date().toUTCString()}`
    };
    var contact1id = (<any>await Xrm.WebApi.createRecord("contact", contact1))
      .id;

    // Create contact 2
    var contact2 = {
      lastname: `Test Contact 2 ${new Date().toUTCString()}`
    };
    var contact2id = (<any>await Xrm.WebApi.createRecord("contact", contact2))
      .id;

    try {
      // Create letter
      const letter1: {
        subject: string;
        letter_activity_parties: {
          participationtypemask: number;
          [something: string]: any;
        }[];
      } = {
        subject: `Sample Letter ${new Date().toUTCString()}`,
        letter_activity_parties: [
          {
            participationtypemask: 2, // To
            "@odata.type": "Microsoft.Dynamics.CRM.activityparty",
            "partyid_contact@odata.bind": `contacts(${contact1id})`
          }
        ]
      };

      var letter1id = (<any>await Xrm.WebApi.createRecord("letter", letter1))
        .id;

      if (!letter1id) throw new Error("Letter not created");

      // Query the letter and check the attribute values
      var letter2 = await Xrm.WebApi.retrieveRecord(
        "letter",
        letter1id,
        "?$expand=letter_activity_parties($select=activitypartyid,_partyid_value,participationtypemask)"
      );

      if (
        !letter2.letter_activity_parties ||
        !letter1.letter_activity_parties.length
      )
        throw new Error("Letter1 letter_activity_parties not set");

      var partyTo: any = findPartyById(
        letter2.letter_activity_parties,
        contact1id
      );

      assert.isNotNull(partyTo, "To Party returned");

      // Add an activity party
      letter1.letter_activity_parties.push({
        participationtypemask: 2, // To
        "@odata.type": "Microsoft.Dynamics.CRM.activityparty",
        "partyid_contact@odata.bind": `contacts(${contact2id})`
      });

      // Update letter
      await Xrm.WebApi.updateRecord("letter", letter1id, letter1);

      // Query the letter and check the attribute values
      var letter3 = await Xrm.WebApi.retrieveRecord(
        "letter",
        letter1id,
        "?$expand=letter_activity_parties($select=activitypartyid,_partyid_value,participationtypemask)"
      );

      var partyTo2: any = findPartyById(
        letter3.letter_activity_parties,
        contact1id
      );
      var partyTo3: any = findPartyById(
        letter3.letter_activity_parties,
        contact2id
      );

      assert.isNotNull(partyTo2, "To Party 1 returned");
      assert.isNotNull(partyTo3, "To Party 2 returned");
    } finally {
      // Delete Contact
      if (contact1id) {
        await Xrm.WebApi.deleteRecord("contact", contact1id);
      }
      if (contact2id) {
        await Xrm.WebApi.deleteRecord("contact", contact2id);
      }
    }
  });

  function findPartyById(parties: { _partyid_value: string }[], id: string) {
    for (let party of parties) {
      if (party._partyid_value == id) {
        return party;
      }
    }
    return null;
  }
});
